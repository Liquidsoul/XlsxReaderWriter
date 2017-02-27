//
//  SwiftTestCase.swift
//  XlsxReaderWriter
//
//  Created by René BIGOT on 07/09/2015.
//  Copyright (c) 2015 BRAE. All rights reserved.
//

import UIKit
import XCTest

enum CellValue {
    case string(String)
    case date(Date)
    case integer(Int)
    case float(Float)
}

extension BRACell {
    func setValue(_ value: CellValue) {
        switch value {
        case let .string(value):
            self.setStringValue(value)
        case let .date(value):
            self.setDateValue(value)
        case let .integer(value):
            self.setIntegerValue(value)
        case let .float(value):
            self.setFloatValue(value)
        }
    }
}

class SwiftTestCase: XCTestCase {

    func testSwiftOpenClose() {
        // This is an example of a functional test case.
        let documentPath = Bundle(for: self.classForCoder).path(forResource: "testWorkbook", ofType: "xlsx")
        NSLog("%@", documentPath!)
        
        let odp: BRAOfficeDocumentPackage = BRAOfficeDocumentPackage.open(documentPath)
        XCTAssertNotNil(odp, "Office document package should not be nil")

        XCTAssertNotNil(odp.workbook, "Office document package should contain a workbook")

        let worksheet: BRAWorksheet = odp.workbook.worksheets[0] as! BRAWorksheet;
        XCTAssertNotNil(worksheet, "Worksheet should not be nil")

        let paths: Array = NSSearchPathForDirectoriesInDomains(FileManager.SearchPathDirectory.documentDirectory, FileManager.SearchPathDomainMask.userDomainMask, true) as Array
        let fullPath: String = paths[0] + "/testSwiftOpenClose.xlsx"
        odp.save(as: fullPath)
        XCTAssert(FileManager.default.fileExists(atPath: fullPath), "No file exists at \(fullPath)")
    }

    func testMyUseCase() {
        let documentPath = Bundle(for: self.classForCoder).path(forResource: "testWorkbook", ofType: "xlsx")
        let odp: BRAOfficeDocumentPackage = BRAOfficeDocumentPackage.open(documentPath)
        guard let worksheet = odp.workbook.worksheets.first as? BRAWorksheet else {
            XCTFail()
            return
        }

        var templateCells = [String: BRACell]()

        for cell in worksheet.cells {
            guard let cell = cell as? BRACell,
                let strValue = cell.stringValue(),
                strValue.hasPrefix("#"),
                strValue.hasSuffix("#")
                else {
                    continue
            }
            print("cell: \(cell)")
            print("str: \(strValue)")
            templateCells[strValue] = cell
        }
//        for row in worksheet.rows {
//            guard let row = row as? BRARow else {
//                continue
//            }
//            for cell in row.cells {
//                guard let cell = cell as? BRACell,
//                    let strValue = cell.stringValue(),
//                    strValue.hasPrefix("#"),
//                    strValue.hasSuffix("#")
//                else {
//                    continue
//                }
//                print("cell: \(cell)")
//                print("str: \(strValue)")
//                templateCells[strValue] = cell
//            }
//        }

        let mapping: [String: CellValue] = [
            "#CLIENT#": .string("The client"),
            "#DATE#": .date(Date()),
            "#COMPTE#": .float(15.12),
            "#PRIX#": .integer(42)
        ]

        for templateKey in templateCells.keys {
            let cell = templateCells[templateKey]!
            let injectedValue = mapping[templateKey] ?? .string("")
            cell.setValue(injectedValue)
        }

        guard let documentsDirPath = NSSearchPathForDirectoriesInDomains(FileManager.SearchPathDirectory.documentDirectory, FileManager.SearchPathDomainMask.userDomainMask, true).first else {
            XCTFail()
            return
        }
        let fullPath = (documentsDirPath as NSString).appendingPathComponent("testSwiftOpenClose.xlsx")
        NSLog("\(fullPath)")
        odp.save(as: fullPath)
    }
}
