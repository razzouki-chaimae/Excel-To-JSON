package com.idemia.exceltojson.viewModel

import android.net.Uri
import android.util.Log
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.result.ActivityResultLauncher
import androidx.activity.result.contract.ActivityResultContracts
import androidx.lifecycle.LiveData
import androidx.lifecycle.MutableLiveData
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import com.google.gson.GsonBuilder
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.launch
import kotlinx.coroutines.yield
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import kotlin.system.exitProcess


class ExcelToJsonViewModel : ViewModel() {

    // about OptionsDialog
    /*private val _showDialog = MutableStateFlow(false)
    val showDialog = _showDialog.asStateFlow()

    fun showDialog() {
        viewModelScope.launch {
            _showDialog.value = true
        }
    }

    fun dismissDialog() {
        viewModelScope.launch {
            _showDialog.value = false
        }
    }

    fun confirmDialog() {
        viewModelScope.launch {
            // Handle confirm button click here
            _showDialog.value = false
            Log.d("confirmDialog", "Here we can choose the excel file that contains the select options")
        }
    }*/

    private val _showDialog = MutableLiveData<Boolean>(false)
    val showDialog : LiveData<Boolean> = _showDialog

    private val _dialogMessage = MutableLiveData<String>("")
    val dialogMessage : LiveData<String> = _dialogMessage

    fun onDialogConfirmed() {
        // Perform the confirm action
        _showDialog.value = false
    }

    fun onDialogDismissed() {
        _showDialog.value = false
    }

    fun showOptionsDialog(message : String) {
        _showDialog.value = true
        _dialogMessage.value = message
    }

    //about conversion of the excel file into json

    private lateinit var pickExcelFileLauncher: ActivityResultLauncher<String>

    fun initialize(activity: ComponentActivity) {
        // use the contract to allow the user to pick an Excel file from their internal storage
        pickExcelFileLauncher = activity.activityResultRegistry.register(
            "excel_file_picker",
            ActivityResultContracts.GetMultipleContents()
        ) { uris ->
            // Process the selected files
            for (uri in uris) {
                // Handle the selected Excel file URI
                if (uri != null) {
                    Log.d(
                        "registerForActivityResult",
                        "+++++ The selected file path is : ${uri.path} +++++"
                    )
                    // TODO : review the following block of code
                    /*var selectedFileName : String? = "demographics"
                    if(uri.path?.contains("/") == true){
                        selectedFileName = uri.path?.substring(
                            uri.path!!.lastIndexOf("/") + 1,
                            uri.path!!.lastIndexOf(".")
                        )
                    } else {
                        if(uri.path?.contains(".") == true) {
                            selectedFileName = uri.path?.substring(
                                uri.path!!.lastIndexOf(".")
                            )
                        }
                    }
                    selectedFileName = selectedFileName.toString()*/
                    val selectedFileName = uri.path?.substring(
                        uri.path!!.lastIndexOf("/") + 1,
                        uri.path!!.lastIndexOf(".")
                    )
                    Log.d(
                        "registerForActivityResult",
                        "+++++ The selected file is : $selectedFileName +++++"
                    )

                    // Make sure that the selected file is an excel one
                    /*if (!uri.path?.substring(uri.path!!.lastIndexOf(".") + 1).equals("xlsx")) {
                        Log.e(
                            "registerForActivityResult",
                            "+++++ Inconvenient document type! +++++"
                        )
                        Toast.makeText(
                            activity.applicationContext,
                            "Inconvenient document type!",
                            Toast.LENGTH_LONG
                        ).show()
                        exitProcess(-1)
                    }*/

                    // read the selected excel file according to which file is selected
                    // demographics.xlsx , demographics_law_enforcement.xlsx or portrait_config.xlsx
                    Log.d("registerForActivityResult", "+++++ Reading from excel file +++++")
                    val rowsList = readExcelFile(activity, selectedFileName, uri)
                    Log.d("registerForActivityResult", "+++++ Reading finished +++++")
                    // convert the excel file into a json one
                    Log.d(
                        "registerForActivityResult",
                        "+++++ Conversion into Json format ... +++++"
                    )
                    // Convert the list to JSON using the Gson library
                    val json = convertExcelToJson(rowsList)
                    Log.d("registerForActivityResult", "+++++ Conversion successful +++++")
                    Log.d("registerForActivityResult", "+++++ Saving Json file ... +++++")
                    if (json != null) {
                        if (selectedFileName != null)
                            saveJsonFile("$selectedFileName.json", json, activity)
                        else
                        // TODO : Ensure that the default name can be "demographics"
                            saveJsonFile("demographics.json", json, activity)
                    } else {
                        Log.d(
                            "registerForActivityResult",
                            "+++++ Saving Json file failed : issue with the output of the conversion step +++++"
                        )
                    }
                }
            }
        }
    }

    fun chooseFile() {

        //  launch the file picker dialog
        pickExcelFileLauncher.launch("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    }

    // read the selected excel file according to which file is selected
    // demographics.xlsx , demographics_law_enforcement.xlsx or portrait_config.xlsx
    private fun readExcelFile(
        activity: ComponentActivity,
        selectedFileName: String?,
        uri: Uri
    ): MutableMap<String, Any> {
        var fileData = mutableMapOf<String, Any>()
        when (selectedFileName) {
            "demographics" -> fileData = readDemographicsFile(activity, uri)
            "demographics_law_enforcement" -> fileData =
                readDemographicsLowEnforcementFile(activity, uri)
            else -> Log.e("readExcelFileFunction", "Excel file unrecognized")
        }

        return fileData

    }

    private fun readDemographicsFile(
        activity: ComponentActivity,
        uri: Uri
    ): MutableMap<String, Any> {

        // get the contentResolver
        val contentResolver = activity.applicationContext.contentResolver
        val inputStream = contentResolver.openInputStream(uri)

        // Create a POI File System object
        //val myFileSystem = POIFSFileSystem(inputStream)
        // Create a workbook using the File System
        //val workbook = HSSFWorkbook(myFileSystem)
        // OR
        // create XSSFWorkBook object
        val workbook = XSSFWorkbook(inputStream)

        // create a variable that will hold all the sheet's details
        val allFields = mutableListOf<Map<String, Any>>()

        // Loop through all the sheets of the excel file
        val sheetIterator = workbook.sheetIterator()
        viewModelScope.launch {
            while (sheetIterator.hasNext()) {
                val sheet = sheetIterator.next()

                val fieldDetails = mutableMapOf<String, Any>()
                fieldDetails["label"] = sheet.sheetName
                Log.d("readExcelFile function", "+++++ sheet name : ${sheet.sheetName} +++++")

                // Create a list to hold the rows
                val rowsList = mutableListOf<Map<String, Any>>()

                // Loop through the rows in the sheet
                for (i in 1 until sheet.physicalNumberOfRows) {
                    val row = sheet.getRow(i)
                    val rowData = mutableMapOf<String, Any>()
                    // create a list holding all the validations of the current line/element
                    val serverValidations = mutableListOf<Map<String, Any>>()

                    // Loop through the cells in the row
                    for (j in 0 until row.lastCellNum) {

                        val columnName = sheet.getRow(0).getCell(j).stringCellValue

                        // TODO : null cell not handled
                        if(columnName == null)
                            Log.e("readExcelFile function",
                                "++++++ cell is null"
                            )

                        Log.d(
                            "readExcelFile function",
                            "+++++ column name : $columnName"
                        )

                        val cell = row.getCell(j)
                        Log.d("readExcelFile function", "++++++++ cell value : $cell")

                        // we create DataFormatter so we can format numeric values as a string using formatter.formatCellValue(cell)
                        val formatter = DataFormatter()

                        if (cell != null) {
                            // Get the cell value according to its type
                            var cellValue = when (cell.cellType) {
                                CellType.NUMERIC -> formatter.formatCellValue(cell)
                                CellType.BOOLEAN -> cell.booleanCellValue
                                else -> cell.stringCellValue
                            }

                            // convert the string "true" and "false" to boolean
                            if (cellValue.equals("true"))
                                cellValue = cell.stringCellValue.toBoolean()
                            if (cellValue.equals("false"))
                                cellValue = cell.stringCellValue.toBoolean()

                            //print("$cellValue | ")

                            // if this column is a validation one we can add its value to the validation list
                            when(columnName){
                                "required" -> {
                                    if (cellValue == true){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "This field is mandatory",
                                            "localizableName" to "fieldMandatory")
                                        serverValidations.add(validation)
                                    }
                                }
                                "maxLength" -> {
                                    if(cellValue.toString().isNotEmpty()){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "This should not exceed $cellValue characters",
                                            "localizableName" to "shouldNotExceed",
                                            "value" to cellValue)
                                        serverValidations.add(validation)
                                    }
                                }
                                "minLength" -> {
                                    if(cellValue.toString().isNotEmpty()){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "This should not be less than $cellValue characters",
                                            "localizableName" to "shouldNotBeLess",
                                            "value" to cellValue)
                                        serverValidations.add(validation)
                                    }
                                }
                                "pattern" -> {
                                    if(cellValue.toString().isNotEmpty()){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "The entered expression is not valid",
                                            "localizableName" to "expressionNotValid",
                                            "value" to cellValue)
                                        serverValidations.add(validation)
                                    }
                                }
                                "email" -> {
                                    if(cellValue.equals("yes")){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "This field is invalid",
                                            "localizableName" to "fieldInvalid")
                                        serverValidations.add(validation)
                                    }
                                }
                                else -> {
                                    // Add the cell value to the row data
                                    rowData[sheet.getRow(0).getCell(j).stringCellValue] = cellValue
                                }
                            }
                        }
                    }

                    rowData["serverValidations"] = serverValidations
                    //println()
                    // Add the row data to the list
                    rowsList.add(rowData)
                }
                fieldDetails["fieldsConfig"] = rowsList

                allFields.add(fieldDetails)
            }
        }


        val allFileData = mutableMapOf<String, Any>()
        // TODO : alpha-config-1 should be a variable which we get from an other source
        allFileData["_id"] = "alpha-config-1"
        allFileData["fieldsGroups"] = allFields
        allFileData["_class"] = "com.idemia.configuration.manager.model.AlphaConfig"
        return allFileData
    }

    private fun readDemographicsLowEnforcementFile(
        activity: ComponentActivity,
        uri: Uri
    ): MutableMap<String, Any> {

        // get the contentResolver
        val contentResolver = activity.applicationContext.contentResolver
        val inputStream = contentResolver.openInputStream(uri)

        // Create a POI File System object
        //val myFileSystem = POIFSFileSystem(inputStream)
        // Create a workbook using the File System
        //val workbook = HSSFWorkbook(myFileSystem)
        // OR
        // create XSSFWorkBook object
        val workbook = XSSFWorkbook(inputStream)

        // create a variable that will hold all the sheet's details
        val allFields = mutableListOf<Map<String, Any>>()

        // Loop through all the sheets of the excel file
        val sheetIterator = workbook.sheetIterator()
        var mainCoroutine = viewModelScope.launch {
            while (sheetIterator.hasNext()) {
                val sheet = sheetIterator.next()

                val fieldDetails = mutableMapOf<String, Any>()
                fieldDetails["label"] = sheet.sheetName
                Log.d("readExcelFile function", "+++++ sheet name : ${sheet.sheetName} +++++")

                // Create a list to hold the rows
                val rowsList = mutableListOf<Map<String, Any>>()

                // Loop through the rows in the sheet
                for (i in 1 until sheet.physicalNumberOfRows) {
                    val row = sheet.getRow(i)
                    val rowData = mutableMapOf<String, Any>()
                    // create a list holding all the validations of the current line/element
                    val serverValidations = mutableListOf<Map<String, Any>>()

                    // Loop through the cells in the row
                    for (j in 0 until row.lastCellNum) {

                        val columnName = sheet.getRow(0).getCell(j).stringCellValue

                        // TODO : null cell not handled
                        if(columnName == null)
                            Log.e("readExcelFile function",
                                "++++++ cell is null"
                            )

                        Log.d(
                            "readExcelFile function",
                            "+++++ column name : $columnName"
                        )

                        val cell = row.getCell(j)
                        Log.d("readExcelFile function", "++++++++ cell value : $cell")

                        // we create DataFormatter so we can format numeric values as a string using formatter.formatCellValue(cell)
                        val formatter = DataFormatter()

                        if (cell != null) {
                            Log.d("readExcelFile function", "########## cell : $cell")
                            // Get the cell value according to its type
                            var cellValue = when (cell.cellType) {
                                CellType.NUMERIC -> formatter.formatCellValue(cell)
                                CellType.BOOLEAN -> cell.booleanCellValue
                                else -> cell.stringCellValue
                            }

                            // convert the string "true" and "false" to boolean
                            if (cellValue.equals("true"))
                                cellValue = cell.stringCellValue.toBoolean()
                            if (cellValue.equals("false"))
                                cellValue = cell.stringCellValue.toBoolean()

                            //print("$cellValue | ")

                            // if this column is a validation one we can add its value to the validation list
                            when(columnName){
                                "required" -> {
                                    if (cellValue == true){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "This field is mandatory",
                                            "localizableName" to "fieldMandatory")
                                        serverValidations.add(validation)
                                    } else Log.d("readExcelFile function", "/////////// Not Required ////////////")
                                }
                                "maxLength" -> {
                                    if(cellValue.toString().isNotEmpty()){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "This should not exceed $cellValue characters",
                                            "localizableName" to "shouldNotExceed",
                                            "value" to cellValue)
                                        serverValidations.add(validation)
                                    } else Log.d("readExcelFile function", "/////////// No Max Length ////////////")
                                }
                                "minLength" -> {
                                    if(cellValue.toString().isNotEmpty()){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "This should not be less than $cellValue characters",
                                            "localizableName" to "shouldNotBeLess",
                                            "value" to cellValue)
                                        serverValidations.add(validation)
                                    } else Log.d("readExcelFile function", "/////////// No Min Length ////////////")
                                }
                                "pattern" -> {
                                    if(cellValue.toString().isNotEmpty()){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "The entered expression is not valid",
                                            "localizableName" to "expressionNotValid",
                                            "value" to cellValue)
                                        serverValidations.add(validation)
                                    } else Log.d("readExcelFile function", "/////////// No pattern ////////////")
                                }
                                "email" -> {
                                    if(cellValue.equals("yes")){
                                        val validation = mutableMapOf("name" to columnName,
                                            "message" to "This field is invalid",
                                            "localizableName" to "fieldInvalid")
                                        serverValidations.add(validation)
                                    } else Log.d("readExcelFile function", "/////////// Not An Email ////////////")
                                }
                                "type" -> {
                                    if(cellValue.equals("select")){
                                        // Pause the coroutine using yield
                                        yield()
                                        Log.d("readExcelFile function", "***** The coroutine is paused *****")
                                        // ask the user to select the excel file that contains the options of the select
                                        Log.d("readExcelFile function", "***** give me ${row.getCell(2).stringCellValue}'s options")
                                        //activity.OptionsDialog(ExcelToJsonViewModel.this)
                                        showOptionsDialog(row.getCell(2).stringCellValue)
                                        Log.d("readExcelFile function", "****** The dialog is shown *****")
                                    }
                                }
                                /*"options" -> {
                                    if(cellValue!=null) {
                                        // Parse the reference to extract the workbook name, sheet name, and column letter
                                        val regex = "^(.+?)!'(.+?)\\[(.+?)\\]\$".toRegex()
                                        val match = regex.find(cell.numericCellValue.toString())
                                        val (workbookName, sheetName, columnLetter) = match!!.destructured
                                        Log.d("readExcelFile function", "+++++ workbookName : $workbookName")
                                        Log.d("readExcelFile function", "+++++ sheetName : $sheetName")
                                        Log.d("readExcelFile function", "+++++ columnLetter : $columnLetter")

                                        // Open the other workbook and get the sheet and column
                                        val optionsFile = FileInputStream(File(workbookName))
                                        val optionsWorkbook = XSSFWorkbook(optionsFile)
                                        val optionsSheet = optionsWorkbook.getSheet(sheetName)

                                        // create a variable that will contain all the options of this select element
                                        val optionsElements = mutableListOf<String>()

                                        // Loop through the cells in the column and read the values
                                        for (k in optionsSheet.firstRowNum..optionsSheet.lastRowNum) {
                                            val optionsRow = optionsSheet.getRow(k)
                                            val optionsCell = optionsRow.getCell(CellReference.convertColStringToIndex(columnLetter))
                                            optionsElements.add(optionsCell.stringCellValue)
                                        }

                                        // Close the input streams and workbooks
                                        optionsFile.close()
                                        optionsWorkbook.close()

                                        Log.d("readExcelFile function", "optionsElements : $optionsElements")
                                        rowData[columnName] = optionsElements
                                    }
                                }*/
                                else -> {
                                    // Add the cell value to the row data
                                    rowData[columnName] = cellValue
                                }
                            }
                        }
                    }

                    rowData["serverValidations"] = serverValidations
                    //println()
                    // Add the row data to the list
                    rowsList.add(rowData)
                }
                fieldDetails["fieldsConfig"] = rowsList

                allFields.add(fieldDetails)
            }
        }


        val allFileData = mutableMapOf<String, Any>()
        // TODO : alpha-config-1 should be a variable which we get from an other source
        allFileData["_id"] = "alpha-config-6"
        allFileData["fieldsGroups"] = allFields
        return allFileData
    }

    private fun convertExcelToJson(rowsList: MutableMap<String, Any>): String? {

        // Create a Gson object with the method called
        // so that it will not convert the apostrophe to its code
        val gson = GsonBuilder().disableHtmlEscaping().create()

        //Convert the Json object to a string
        val jsonString = gson.toJson(rowsList)

        // Print the string representation of the Json object
        Log.d("convertExcelToJson function", jsonString)

        return jsonString
    }

    private fun saveJsonFile(
        fileName: String = "demographics",
        json: String,
        activity: ComponentActivity
    ) {
        // TODO : ask the user where to store the file

        val documentsFolder =
            activity.baseContext.getExternalFilesDir(null) // get the documents folder path
        Log.d("saveJsonFile function", "+++++ storage directory : $documentsFolder +++++")
        val file = File(documentsFolder, fileName) // create a new file with the specified file name
        file.writeText(json) // write the JSON string to the file
        Log.d("saveJsonFile function", "+++++ Saving successful +++++")
    }

}