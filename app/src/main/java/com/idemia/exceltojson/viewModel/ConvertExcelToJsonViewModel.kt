package com.idemia.exceltojson.viewModel

import android.annotation.SuppressLint
import android.app.Activity
import android.app.Application
import android.content.Intent
import android.net.Uri
import android.os.Environment
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.core.app.ActivityCompat.startActivityForResult
import androidx.lifecycle.ViewModel
import androidx.lifecycle.ViewModelProvider
import com.google.gson.GsonBuilder
//import com.idemia.exceltojson.model.ConverTExcelToJsonContract
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.BufferedWriter
import java.io.File
import java.io.FileWriter
import kotlin.system.exitProcess

class ConvertExcelToJsonViewModel (application: Application) : ViewModelProvider.Factory {

    override fun <T : ViewModel> create(modelClass: Class<T>): T {
        TODO("Not yet implemented")
        //if (modelClass.isAssignableFrom(ConvertExcelToJsonViewModel::class.java)) {
        //    return ConvertExcelToJsonViewModel(application) as T
        //}
        //throw IllegalArgumentException("Unknown ViewModel class")
    }

    // create a variable to hold the application context
    private val appContext = application

    private val context = appContext.applicationContext

    // Request code for selecting a PDF document.
    private val requestCodePickExcelFile = 1

    //private val conversionContract = ConverTExcelToJsonContract()

    private val _uri = MutableStateFlow<Uri?>(null)
    val uri : StateFlow<Uri?> = _uri

    //private val conversionLauncher = registerForActivityResult(ActivityResultContracts.GetContent()) { uri -> _uri.value = uri}

    fun onClickUploadButton(fileUri: Uri) {
        convertExcelIntoJson(fileUri)
    }


    private fun convertExcelIntoJson(fileUri: Uri) {
        //chooseFile()

        // read the selected excel file
        println("+++++++++++++++++ Reading from excel file +++++++++++++++++")
        val rowsList = readExcelFile(fileUri)
        println("+++++++++++++++++ Reading finished +++++++++++++++++")
        // convert the excel file into a json one
        println("+++++++++++++++++ Conversion into Json format ... +++++++++++++++++")
        // Convert the list to JSON using the Gson library
        val json = convertExcelToJson(rowsList)
        // Close the workbook and input stream
        //workbook.close()
        //inputStream?.close()
        println("+++++++++++++++++ Conversion successful +++++++++++++++++")
        println("+++++++++++++++++ Saving Json file ... +++++++++++++++++")
        if (json != null) {
            saveJsonFile(json)
        }else {
            println("+++++++++++++++++ Saving Json file failed : issue with the output of the conversion +++++++++++++++++")
        }
    }

    //private fun chooseFile() {

    //    conversionContract.launch()
    //    conversionLauncher.launch("application/vnd.ms-excel")

    //    val pickExcelFileLauncher =
    //        registerForActivityResult(ActivityResultContracts.CreateDocument()) { uri ->
                // handle the picked file URI
    //            if (uri != null) {
                    // do something with the picked file URI
    //            }
    //        }

        //ask the user to specify the file he want
    //    val intent = Intent(Intent.ACTION_OPEN_DOCUMENT).apply {
    //        addCategory(Intent.CATEGORY_OPENABLE)
    //        type = "application/*"
            //type = "application/vnd.ms-excel"

            // Optionally, specify a URI for the file that should appear in the
            // system file picker when it loads.
            //putExtra(DocumentsContract.EXTRA_INITIAL_URI, pickerInitialUri)
    //    }
    //    startActivityForResult(intent, requestCodePickExcelFile)
    //}

    //override fun onActivityResult(requestCode: Int, resultCode: Int, resultData: Intent?) {
    //    super.onActivityResult(requestCode, resultCode, resultData)
    //    if (requestCode == requestCodePickExcelFile && resultCode == Activity.RESULT_OK) {
            // The result data contains a URI for the document or directory that
            // the user selected.
    //        val uri = resultData?.data
    //        if (uri != null) {
    //            println("++++++++++++++ The selected file is : ${uri.path} +++++++++++++++")
                // Make sure that the selected file is an excel one
    //            if (!uri.path?.substring(uri.path!!.lastIndexOf(".") + 1).equals("xlsx")) {
    //                Toast.makeText(this, "Inconvenient document type!", Toast.LENGTH_LONG).show()
    //                exitProcess(-1)
    //            }
                // read the selected excel file
    //            println("+++++++++++++++++ Reading from excel file +++++++++++++++++")
    //            val rowsList = readExcelFile(uri)
    //            println("+++++++++++++++++ Reading finished +++++++++++++++++")
                // convert the excel file into a json one
    //            println("+++++++++++++++++ Conversion into Json format ... +++++++++++++++++")
                // Convert the list to JSON using the Gson library
    //            val json = convertExcelToJson(rowsList)
                // Close the workbook and input stream
                //workbook.close()
                //inputStream?.close()
    //            println("+++++++++++++++++ Conversion successful +++++++++++++++++")
    //            println("+++++++++++++++++ Saving Json file ... +++++++++++++++++")
    //            if (json != null) {
    //                saveJsonFile(json)
    //            } else {
    //                println("+++++++++++++++++ Saving Json file failed : issue with the output of the conversion +++++++++++++++++")
    //            }
    //        }
    //    }
    //}

    private fun readExcelFile(uri: Uri): MutableList<Map<String, Any>> {
        // get the contentResolver
        val contentResolver = appContext.contentResolver
        val inputStream = contentResolver.openInputStream(uri)

        // Create a list to hold the rows
        val rowsList = mutableListOf<Map<String, Any>>()

        // Create a POI File System object
        //val myFileSystem = POIFSFileSystem(inputStream)
        // Create a workbook using the File System
        //val workbook = HSSFWorkbook(myFileSystem)
        // OR
        // create XSSFWorkBook object
        val workbook = XSSFWorkbook(inputStream)

        // TODO : fetch all the sheets of the excel file
        // access to the first sheet in the excel file
        //val sheet = workbook.getSheet("Applicant Details")

        // Loop through all the sheets of the excel file
        val sheetIterator = workbook.sheetIterator()
        while (sheetIterator.hasNext()) {
            val sheet = sheetIterator.next()

            println("++++++++++++ sheet name : ${sheet.sheetName} ++++++++++++++")
            // Loop through the rows in the sheet
            for (i in 1 until sheet.physicalNumberOfRows) {
                val row = sheet.getRow(i)
                val rowData = mutableMapOf<String, Any>()

                // Loop through the cells in the row
                for (j in 0 until row.physicalNumberOfCells) {
                    val cell = row.getCell(j)

                    // Get the cell value as a string
                    val cellValue = when (cell.cellType) {
                        CellType.NUMERIC -> cell.numericCellValue
                        CellType.BOOLEAN -> cell.booleanCellValue
                        else -> cell.stringCellValue
                    }
                    print("$cellValue | ")

                    // Add the cell value to the row data
                    rowData[sheet.getRow(0).getCell(j).stringCellValue] = cellValue
                }

                println()
                // Add the row data to the list
                rowsList.add(rowData)
            }
        }
        return rowsList
    }

    private fun convertExcelToJson(rowsList: MutableList<Map<String, Any>>): String? {
        val gson = GsonBuilder().setPrettyPrinting().create()
        val json = gson.toJson(rowsList)

        // Print the JSON output
        println(json)

        return json
    }

    private fun saveJsonFile(json: String) {
        // TODO : ask the user where to store the file
        //Solution 3
        val storageDir = context.getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS)
        println("+++++++++++++++ storage directory : $storageDir ++++++++++++++++++++")
        if (storageDir != null) {
            println("++++++++++++++ storage directory is not null +++++++++++++++")
            if (!storageDir.exists()) {
                println("++++++++++++++ storage directory does not exist +++++++++++++++")
                storageDir.mkdir()
                println("++++++++++++++ storage directory made +++++++++++++++")
            }
            val file = File.createTempFile(
                "ExcelToJson",
                ".json",
                storageDir
            )
            val output = BufferedWriter(FileWriter(file))
            output.write(json)
            output.close()
            println("+++++++++++++++++ Saving successful +++++++++++++++++")
        } else
            println("++++++++++++++ storage directory is null +++++++++++++++")
    }



}