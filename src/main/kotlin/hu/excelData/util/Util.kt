package hu.excelData.util

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream

public fun writeWorkBook(workbook: XSSFWorkbook, fileName: String) {
    val out = FileOutputStream(File(fileName))
    workbook.write(out)
    out.close()
    println("$fileName written successfully on disk.")


}

