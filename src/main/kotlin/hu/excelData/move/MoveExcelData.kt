package hu.excelData.move

import hu.excelData.util.writeWorkBook
import java.io.File
import java.io.FileInputStream
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.util.*

public class MoveExcelData() {


    val idColumn = 0
    val szallasColumn = 1
    val szallasRow = 10
    val vendegLatasColumn = 13
    val vendegLatasRow =11
    val szabadIdoColumn = 25
    val szabadIdoRow = 12

    public fun moveExcelData(dataRootDirectory: String) {

        val workbook = XSSFWorkbook();
        val createHelper = workbook.getCreationHelper();

        //Create a blank sheet
        val summerizeSheet = initSummerizeSheet(workbook)

        File(dataRootDirectory).listFiles().filter { file -> file.name.endsWith("xls") || file.name.endsWith("xlsx") }.forEachIndexed { i,  file ->
            val sheet = readXLS(file)
            val id = readID(sheet)
            val rowNum = i +1
            val row = summerizeSheet.createRow(rowNum)
            row.createCell(0).setCellValue(id.toDouble())
            val szallasData = extractDataFromSheet(sheet,szallasRow)
            writeDataIntoSheet(summerizeSheet,szallasData,rowNum,szallasColumn)
            val vendegLatasData = extractDataFromSheet(sheet,vendegLatasRow)
            writeDataIntoSheet(summerizeSheet,vendegLatasData,rowNum,vendegLatasColumn)
            val szabadIdoData = extractDataFromSheet(sheet,szabadIdoRow)
            writeDataIntoSheet(summerizeSheet,szabadIdoData,rowNum,szabadIdoColumn)
        }

        writeWorkBook(workbook,"./data/summarizeData.xlsx")

    }

    public fun initSummerizeSheet(workbook : XSSFWorkbook) : XSSFSheet{
       val sheet = workbook.createSheet("Employee Data");
        val row =sheet.createRow(0)
        row.createCell(1).setCellValue("Szallashely")
        sheet.addMergedRegion(CellRangeAddress(0, 0, 1, 12))
        row.createCell(13).setCellValue("Vendeglatas")
        sheet.addMergedRegion(CellRangeAddress(0, 0, 13, 24))
        row.createCell(25).setCellValue("Szabadido")
        sheet.addMergedRegion(CellRangeAddress(0, 0, 25, 36))
        return sheet
    }

    public fun extractDataFromSheet(sheet: XSSFSheet, rowNumber : Int) : List<Double>{
        val row = sheet.getRow(rowNumber)
        val result = LinkedList<Double>()

        for (i in 2..13){
            println(row.getCell(i))
            if (row.getCell(i).cellType == 0){
                result.add(row.getCell(i).numericCellValue)
            } else {
                result.add(0.0)
            }
        }

        return result
    }

    public fun writeDataIntoSheet(sheet: XSSFSheet,data : List<Double>,rowNumber: Int, columnNumber : Int){

        val row = sheet.getRow(rowNumber)
        data.forEachIndexed { i, data ->
            row.createCell(i + columnNumber).setCellValue(data)
        }
    }

    public fun readXLS(file: File): XSSFSheet {

        //Get the workbook instance for XLS file
        val workbook = XSSFWorkbook(FileInputStream(file))//HSSFWorkbook(FileInputStream(file))

        //Get first sheet from the workbook
        val sheet = workbook.getSheetAt(0)
        val row = sheet.getRow(12)

        /*for (i in  0..25) {
            println("$i\t${row.getCell(i)}")
        }*/
        println("$file wes read...")
        return sheet
    }

    public fun readID(sheet: XSSFSheet) : Long{

        val row = sheet.getRow(2)
        return row.getCell(14).numericCellValue.toLong()
    }

}

fun main(args: Array<String>) {
    val moveExcelData = MoveExcelData()
    //moveExcelData.readXLS(File("./data/xls/Kitöltött_Cafeteria kalkulátor_2016_21001358.xlsx"))
    moveExcelData.moveExcelData("./data/xls/")
}
