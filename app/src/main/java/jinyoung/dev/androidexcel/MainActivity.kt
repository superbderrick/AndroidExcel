package jinyoung.dev.androidexcel

import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.*
import android.widget.TextView

class MainActivity : AppCompatActivity() {
    private lateinit var textView: TextView

    companion object {
        private val SHEET_NAME = "testSheet"
        private val EXCEL_FILE_NAME = "test.xlsx"
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        textView = findViewById(R.id.textView)

        //Basic stuffs
        val testWorkBook = createWorkbook()
        createExcelFile(testWorkBook)

    }
    private fun createWorkbook(): Workbook {
        // Creating a workbook object from the XSSFWorkbook() class
        val ourWorkbook = XSSFWorkbook()

        //Creating a sheet called "statSheet" inside the workbook and then add data to it
        val sheet: Sheet = ourWorkbook.createSheet(SHEET_NAME)
        addData(sheet)

        return ourWorkbook
    }

    private fun addData(sheet: Sheet) {

        //Creating rows at passed in indices
        val row1 = sheet.createRow(0)
        val row2 = sheet.createRow(1)
        val row3 = sheet.createRow(2)


        //Adding data to each  cell
        createCell(row1, 0, "Name")
        createCell(row1, 1, "Derrick")

        createCell(row2, 0, "AGE")
        createCell(row2, 1, "34")

        createCell(row3, 0, "JOB")
        createCell(row3, 1, "SWE")

    }

    //function for creating a cell.
    private fun createCell(sheetRow: Row, columnIndex: Int, cellValue: String?) {
        //create a cell at a passed in index
        val ourCell = sheetRow.createCell(columnIndex)
        //add the value to it
        ourCell?.setCellValue(cellValue)
    }

    private fun createExcelFile(ourWorkbook: Workbook) {

        //get our app file directory
        val ourAppFileDirectory = filesDir
        //Check whether it exists or not, and create if does not exist.
        if (ourAppFileDirectory != null && !ourAppFileDirectory.exists()) {
            ourAppFileDirectory.mkdirs()
        }

        //Create an excel file called test.xlsx
        val excelFile = File(ourAppFileDirectory, EXCEL_FILE_NAME)

        //Write a workbook to the file using a file outputstream
        try {
            val fileOut = FileOutputStream(excelFile)
            ourWorkbook.write(fileOut)
            fileOut.close()
        } catch (e: FileNotFoundException) {
            e.printStackTrace()
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }


}