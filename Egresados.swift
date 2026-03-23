//Da la cantidad de los diferentes egresados de la generacion 2025
//Licenciatura, propedeutico, preparatoria,maestria y doctorado 

import Foundation
import ZIPFoundation

// MARK: - Helpers (XLSX parsing via ZIP + XML)
enum XLSXError: Error {
    case entryNotFound(String)
    case xmlParseFailed(String)
    case sheetNotFound(String)
    case relationshipNotFound(String)
}

final class SharedStringsParser: NSObject, XMLParserDelegate {
    private(set) var shared: [String] = []
    private var currentSI = ""
    private var inT = false

    func parser(_ parser: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName qName: String?, attributes: [String : String] = [:]) {
        if name == "si" {
            currentSI = ""
        } else if name == "t" {
            inT = true
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inT {
            currentSI += string
        }
    }

    func parser(_ parser: XMLParser, didEndElement name: String, namespaceURI: String?, qualifiedName qName: String?) {
        if name == "t" {
            inT = false
        } else if name == "si" {
            shared.append(currentSI)
        }
    }
}

final class WorkbookSheetParser: NSObject, XMLParserDelegate {
    let targetSheetName: String
    private(set) var targetRid: String?

    init(targetSheetName: String) {
        self.targetSheetName = targetSheetName
    }

    private func extractRid(from attributes: [String: String]) -> String? {
        // Common cases: "r:id" or "{.../relationships}id"
        if let v = attributes["r:id"] { return v }
        if let v = attributes.values.first(where: { _ in true }) {
            // not used; kept for safety
            _ = v
        }
        for (k, v) in attributes {
            if k.lowercased().hasSuffix(":id") && k.lowercased().contains("r") {
                return v
            }
        }
        if let v = attributes["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"] {
            return v
        }
        return nil
    }

    func parser(_ parser: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName qName: String?, attributes: [String : String] = [:]) {
        guard name == "sheet" else { return }
        guard attributes["name"]?.trimmingCharacters(in: .whitespacesAndNewlines) == targetSheetName else { return }

        targetRid = extractRid(from: attributes)
    }
}

final class WorkbookRelsParser: NSObject, XMLParserDelegate {
    let targetRid: String
    private(set) var target: String?

    init(targetRid: String) {
        self.targetRid = targetRid
    }

    func parser(_ parser: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName qName: String?, attributes: [String : String] = [:]) {
        guard name == "Relationship" else { return }
        if attributes["Id"] == targetRid {
            target = attributes["Target"]
        }
    }
}

final class ColumnCCountParser: NSObject, XMLParserDelegate {
    private let columnPrefix: String
    private let sharedStrings: [String]
    private let headerToExcludeUpper: String

    private(set) var counts: [String: Int] = [:]

    private var currentRef: String?
    private var currentType: String? // "s" | "inlineStr" | nil
    private var cellV = ""
    private var inlineText = ""

    private var capturingV = false
    private var inInlineIs = false
    private var capturingInlineT = false

    init(columnPrefix: String, sharedStrings: [String], headerToExcludeUpper: String) {
        self.columnPrefix = columnPrefix
        self.sharedStrings = sharedStrings
        self.headerToExcludeUpper = headerToExcludeUpper
    }

    private func isTargetCellRef(_ ref: String?) -> Bool {
        guard let ref else { return false }
        guard ref.hasPrefix(columnPrefix) else { return false }
        // Ensure it looks like C<number>
        let suffix = String(ref.dropFirst(columnPrefix.count))
        return !suffix.isEmpty && suffix.allSatisfy({ $0.isNumber })
    }

    private func normalize(_ s: String) -> String {
        s.trimmingCharacters(in: .whitespacesAndNewlines)
    }

    func parser(_ parser: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName qName: String?, attributes: [String : String] = [:]) {
        if name == "c" {
            currentRef = attributes["r"]
            currentType = attributes["t"] // optional
            cellV = ""
            inlineText = ""
            capturingV = false
            inInlineIs = false
            capturingInlineT = false
        } else if name == "v" {
            if isTargetCellRef(currentRef) && currentType != "inlineStr" {
                capturingV = true
            }
        } else if name == "is" {
            inInlineIs = (isTargetCellRef(currentRef) && currentType == "inlineStr")
        } else if name == "t" {
            if inInlineIs {
                capturingInlineT = true
            }
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if capturingV {
            cellV += string
        } else if capturingInlineT {
            inlineText += string
        }
    }

    func parser(_ parser: XMLParser, didEndElement name: String, namespaceURI: String?, qualifiedName qName: String?) {
        if name == "v" {
            capturingV = false
        } else if name == "is" {
            inInlineIs = false
        } else if name == "t" {
            capturingInlineT = false
        } else if name == "c" {
            guard isTargetCellRef(currentRef) else { return }

            let rawValue: String
            if currentType == "s" {
                let idxStr = normalize(cellV)
                if let idx = Int(idxStr), idx >= 0, idx < sharedStrings.count {
                    rawValue = sharedStrings[idx]
                } else {
                    return
                }
            } else if currentType == "inlineStr" {
                rawValue = inlineText
            } else {
                rawValue = cellV
            }

            let value = normalize(rawValue)
            guard !value.isEmpty else { return }

            let up = value.uppercased()
            if up == headerToExcludeUpper { return }
            if up == "NAN" || up == "NONE" { return }

            counts[value, default: 0] += 1

            // reset
            currentRef = nil
            currentType = nil
            cellV = ""
            inlineText = ""
        }
    }
}

// MARK: - Main
let args = CommandLine.arguments
let xlsxPath = args.dropFirst().first ?? "/home/dulce/Swift/BD_anuario_2025 .xlsx"

let targetSheetName = "Egresados 2025"
let columnPrefix = "C"
let headerToExclude = "NIVEL DE EGRESO"

let fileURL = URL(fileURLWithPath: xlsxPath)
let archive = try Archive(url: fileURL, accessMode: .read)

func readZipEntry(_ entryPath: String) throws -> Data {
    guard let entry = archive[entryPath] else {
        throw XLSXError.entryNotFound(entryPath)
    }
    return try archive.extract(entry)
}

func parseXML(_ data: Data, parserDelegate: NSObject & XMLParserDelegate) throws {
    guard let xmlParser = XMLParser(data: data) else {
        throw XLSXError.xmlParseFailed("XMLParser init failed for data")
    }
    xmlParser.delegate = parserDelegate
    if !xmlParser.parse() {
        throw XLSXError.xmlParseFailed("XML parse failed")
    }
}

// 1) Locate sheet r:id in xl/workbook.xml
let workbookData = try readZipEntry("xl/workbook.xml")
let sheetParser = WorkbookSheetParser(targetSheetName: targetSheetName)
try parseXML(workbookData, parserDelegate: sheetParser)

guard let rid = sheetParser.targetRid else {
    throw XLSXError.sheetNotFound("No se encontró hoja: \(targetSheetName)")
}

// 2) Resolve rid -> Target in xl/_rels/workbook.xml.rels
let relsData = try readZipEntry("xl/_rels/workbook.xml.rels")
let relsParser = WorkbookRelsParser(targetRid: rid)
try parseXML(relsData, parserDelegate: relsParser)

guard let target = relsParser.target else {
    throw XLSXError.relationshipNotFound(rid)
}

let worksheetPath = target.hasPrefix("xl/") ? target : "xl/" + target

// 3) Load shared strings (needed for t="s")
let sharedStrings: [String]
if archive.contains(entry: "xl/sharedStrings.xml") {
    let ssData = try readZipEntry("xl/sharedStrings.xml")
    let ssParser = SharedStringsParser()
    try parseXML(ssData, parserDelegate: ssParser)
    sharedStrings = ssParser.shared
} else {
    sharedStrings = []
}

// 4) Parse worksheet and count column C
let sheetData = try readZipEntry(worksheetPath)
let colParser = ColumnCCountParser(
    columnPrefix: columnPrefix,
    sharedStrings: sharedStrings,
    headerToExcludeUpper: headerToExclude.uppercased()
)
try parseXML(sheetData, parserDelegate: colParser)

// 5) Print results sorted by count desc
let sorted = colParser.counts.sorted { a, b in
    if a.value != b.value { return a.value > b.value }
    return a.key < b.key
}

for (value, count) in sorted {
    print("\(count)\t\(value)")
}
}

// 1) Locate sheet r:id in xl/workbook.xml
let workbookData = try readZipEntry("xl/workbook.xml")
let sheetParser = WorkbookSheetParser(targetSheetName: targetSheetName)
try parseXML(workbookData, parserDelegate: sheetParser)

guard let rid = sheetParser.targetRid else {
    throw XLSXError.sheetNotFound("No se encontró hoja: \(targetSheetName)")
}

// 2) Resolve rid -> Target in xl/_rels/workbook.xml.rels
let relsData = try readZipEntry("xl/_rels/workbook.xml.rels")
let relsParser = WorkbookRelsParser(targetRid: rid)
try parseXML(relsData, parserDelegate: relsParser)

guard let target = relsParser.target else {
    throw XLSXError.relationshipNotFound(rid)
}

let worksheetPath = target.hasPrefix("xl/") ? target : "xl/" + target

// 3) Load shared strings (needed for t="s")
let sharedStrings: [String]
if archive.contains(entry: "xl/sharedStrings.xml") {
    let ssData = try readZipEntry("xl/sharedStrings.xml")
    let ssParser = SharedStringsParser()
    try parseXML(ssData, parserDelegate: ssParser)
    sharedStrings = ssParser.shared
} else {
    sharedStrings = []
}

// 4) Parse worksheet and count column C
let sheetData = try readZipEntry(worksheetPath)
let colParser = ColumnCCountParser(
    columnPrefix: columnPrefix,
    sharedStrings: sharedStrings,
    headerToExcludeUpper: headerToExclude.uppercased()
)
try parseXML(sheetData, parserDelegate: colParser)

// 5) Print results sorted by count desc
let sorted = colParser.counts.sorted { a, b in
    if a.value != b.value { return a.value > b.value }
    return a.key < b.key
}

for (value, count) in sorted {
    print("\(count)\t\(value)")
}
