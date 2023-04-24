package com.idemia.exceltojson.viewModel

import android.net.Uri
import android.util.Log
import androidx.activity.ComponentActivity
import androidx.activity.result.ActivityResultLauncher
import androidx.activity.result.contract.ActivityResultContracts
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import com.google.gson.GsonBuilder
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.w3c.dom.Document
import org.w3c.dom.Element
import java.io.File
import javax.xml.parsers.DocumentBuilderFactory
import javax.xml.transform.TransformerFactory
import javax.xml.transform.dom.DOMSource
import javax.xml.transform.stream.StreamResult

class ExcelToJsonViewModel : ViewModel() {

    private lateinit var pickExcelFileLauncher: ActivityResultLauncher<String>

    fun initialize(activity: ComponentActivity) {
        // use the contract to allow the user to pick an Excel file from their internal storage
        pickExcelFileLauncher = activity.activityResultRegistry.register(
            "excel_file_picker",
            ActivityResultContracts.GetMultipleContents(),
        ) { uris ->

            // Create xml files for the internationalisation
            val frenchXmlFile =
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument()
            val englishXmlFile =
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument()
            val portugueseXmlFile =
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument()
            // Create the root element of each xml doc
            val frenchXmlFileRoot = frenchXmlFile.createElement("resources")
            frenchXmlFile.appendChild(frenchXmlFileRoot)
            val englishXmlFileRoot = englishXmlFile.createElement("resources")
            englishXmlFile.appendChild(englishXmlFileRoot)
            val portugueseXmlFileRoot = portugueseXmlFile.createElement("resources")
            portugueseXmlFile.appendChild(portugueseXmlFileRoot)

            val xmlFiles = mutableListOf<Pair<Document, Element>>()
            xmlFiles.add(Pair<Document, Element>(frenchXmlFile, frenchXmlFileRoot))
            xmlFiles.add(Pair<Document, Element>(englishXmlFile, englishXmlFileRoot))
            xmlFiles.add(Pair<Document, Element>(portugueseXmlFile, portugueseXmlFileRoot))

            // Process the selected files
            for (uri in uris) {
                // Handle the selected Excel file URI
                if (uri != null) {
                    Log.d(
                        "VM > initialize function > registerForActivityResult",
                        "+++++ The selected file path is : ${uri.path} +++++",
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
                        uri.path!!.lastIndexOf("."),
                    )
                    Log.d(
                        "VM > initialize function > registerForActivityResult",
                        "+++++ The selected file is : $selectedFileName +++++",
                    )

                    // read the selected excel file according to which file is selected
                    // demographics.xlsx or demographics_law_enforcement.xlsx
                    Log.d(
                        "VM > initialize function > registerForActivityResult",
                        "+++++ Reading from excel file +++++",
                    )
                    val rowsList = readExcelFile(activity, selectedFileName, uri, xmlFiles)
                    Log.d(
                        "VM > initialize function > registerForActivityResult",
                        "+++++ Reading finished +++++",
                    )
                    // convert the excel file into a json one
                    Log.d(
                        "VM > initialize function > registerForActivityResult",
                        "+++++ Conversion into Json format ... +++++",
                    )
                    // Convert the list to JSON using the Gson library
                    val json = convertExcelToJson(rowsList)
                    Log.d(
                        "VM > initialize function > registerForActivityResult",
                        "+++++ Conversion successful +++++",
                    )
                    Log.d(
                        "VM > initialize function > registerForActivityResult",
                        "+++++ Saving Json file ... +++++",
                    )
                    if (json != null) {
                        if (selectedFileName != null) {
                            saveJsonFile("$selectedFileName.json", json, activity)
                        } else {
                            // TODO : Ensure that the default name can be "demographics"
                            saveJsonFile("demographics.json", json, activity)
                        }
                    } else {
                        Log.d(
                            "VM > initialize function > registerForActivityResult",
                            "+++++ Saving Json file failed : issue with the output of the conversion step +++++",
                        )
                    }
                }
            }
            // Write the XML documents to a files
            val transformer = TransformerFactory.newInstance().newTransformer()

            val documentsFolder =
                activity.baseContext.getExternalFilesDir(null) // get the documents folder path

            val frenchXmlFinalFile = File(
                documentsFolder,
                "french_strings.xml",
            ) // create a new file with the specified file name
            transformer.transform(DOMSource(frenchXmlFile), StreamResult(frenchXmlFinalFile))
            val frenchXmlFinalFileContent = frenchXmlFinalFile.readText() // for debugging
            Log.d(
                "Internationalisation",
                "Generated French XML file is : $frenchXmlFinalFileContent",
            ) // for debugging

            val englishXmlFinalFile = File(documentsFolder, "english_strings.xml")
            transformer.transform(DOMSource(englishXmlFile), StreamResult(englishXmlFinalFile))
            val englishXmlFinalFileContent = englishXmlFinalFile.readText() // for debugging
            Log.d(
                "Internationalisation",
                "Generated English XML file is : $englishXmlFinalFileContent",
            ) // for debugging

            val portugueseXmlFinalFile = File(documentsFolder, "portuguese_strings.xml")
            transformer.transform(
                DOMSource(portugueseXmlFile),
                StreamResult(portugueseXmlFinalFile),
            )
            val portugueseXmlFinalFileContent = portugueseXmlFinalFile.readText() // for debugging
            Log.d(
                "Internationalisation",
                "Generated Portuguese XML file is : $portugueseXmlFinalFileContent",
            ) // for debugging
        }
    }

    fun chooseFile() {
        //  launch the file picker dialog
        pickExcelFileLauncher.launch("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    }

    // read the selected excel file according to which file is selected
    // demographics.xlsx or demographics_law_enforcement.xlsx
    private fun readExcelFile(
        activity: ComponentActivity,
        selectedFileName: String?,
        uri: Uri,
        xmlFiles: MutableList<Pair<Document, Element>>,
    ): MutableMap<String, Any> {
        var fileData = mutableMapOf<String, Any>()
        when (selectedFileName) {
            "demographics" -> fileData = readDemographicsFile(activity, uri, xmlFiles)
            "demographics_law_enforcement" ->
                fileData =
                    readDemographicsLowEnforcementFile(activity, uri, xmlFiles)
            else -> Log.e("VM > readExcelFile function", "Excel file unrecognized")
        }

        return fileData
    }

    private fun readDemographicsFile(
        activity: ComponentActivity,
        uri: Uri,
        xmlFiles: MutableList<Pair<Document, Element>>,
    ): MutableMap<String, Any> {
        // get the contentResolver
        val contentResolver = activity.applicationContext.contentResolver
        val inputStream = contentResolver.openInputStream(uri)

        // Create a POI File System object
        // val myFileSystem = POIFSFileSystem(inputStream)
        // Create a workbook using the File System
        // val workbook = HSSFWorkbook(myFileSystem)
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

                // if it is a sheet that contains only dropdown options we will not read it right now
                // because we read it only on-demand
                if (sheet.sheetName.contains("options")) {
                    continue
                }

                val fieldDetails = mutableMapOf<String, Any>()
                fieldDetails["label"] = sheet.sheetName
                Log.d("readExcelFile function", "+++++ sheet name : ${sheet.sheetName} +++++")

                // set the localizableName of this section (sheet)
                // by removing spaces from the label of this section (sheet)
                var sectionLocalizableName = sheet.sheetName.replace("\\s".toRegex(), "")
                // remove 's from the label of this section (sheet)
                sectionLocalizableName = sectionLocalizableName.replace("'s", "")
                // Now we get the right localizableName, so we can add it
                fieldDetails["localizableName"] = sectionLocalizableName

                // Create a list to hold the rows
                val rowsList = mutableListOf<Map<String, Any>>()

                // Loop through the rows in the sheet
                for (i in 1 until sheet.physicalNumberOfRows) {
                    // Create a variable that will contain the localizableName of this field
                    var localizableName = ""

                    val row = sheet.getRow(i)
                    val rowData = mutableMapOf<String, Any>()
                    // create a list holding all the validations of the current line/element
                    val serverValidations = mutableListOf<Map<String, Any>>()

                    // Loop through the cells in the row
                    for (j in 0 until row.lastCellNum) {
                        val columnName = sheet.getRow(0).getCell(j).stringCellValue

                        // TODO : null cell not handled
                        if (columnName == null) {
                            Log.e(
                                "readExcelFile function",
                                "++++++ cell is null",
                            )
                        }

                        Log.d(
                            "readExcelFile function",
                            "+++++ column name : $columnName",
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
                            if (cellValue.equals("true")) {
                                cellValue = cell.stringCellValue.toBoolean()
                            }
                            if (cellValue.equals("false")) {
                                cellValue = cell.stringCellValue.toBoolean()
                            }

                            // Save the localizableName of this line
                            // so we can use it; if it requires options for the dropdown
                            if (columnName.equals("localizableName")) {
                                localizableName = cell.stringCellValue
                            }

                            // if this column is for internationalisation we should add its value to string.xml
                            // if this column is a validation one we should add its value to the validation list
                            // if the type of this field is select (dropdown) we should read the sheet that contains its options (using its localizableName)
                            when (columnName) {
                                "Français" -> {
                                    // get french file
                                    val (frenchXmlFile, frenchXmlFileRoot) = xmlFiles[0]
                                    // Add a string element
                                    val string = frenchXmlFile.createElement("string")
                                    string.setAttribute("name", localizableName)
                                    string.appendChild(frenchXmlFile.createTextNode(cellValue as String?))
                                    frenchXmlFileRoot.appendChild(string)
                                }
                                "English" -> {
                                    // get english file
                                    val (englishXmlFile, englishXmlFileRoot) = xmlFiles[1]
                                    // Add a string element
                                    val string = englishXmlFile.createElement("string")
                                    string.setAttribute("name", localizableName)
                                    string.appendChild(englishXmlFile.createTextNode(cellValue as String?))
                                    englishXmlFileRoot.appendChild(string)
                                }
                                "Portugais" -> {
                                    // get portuguese file
                                    val (portugueseXmlFile, portugueseXmlFileRoot) = xmlFiles[2]
                                    // Add a string element
                                    val string = portugueseXmlFile.createElement("string")
                                    string.setAttribute("name", localizableName)
                                    string.appendChild(portugueseXmlFile.createTextNode(cellValue as String?))
                                    portugueseXmlFileRoot.appendChild(string)
                                }
                                "required" -> {
                                    if (cellValue == true) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "This field is mandatory",
                                            "localizableName" to "fieldMandatory",
                                        )
                                        serverValidations.add(validation)
                                    }
                                }
                                "maxLength" -> {
                                    if (cellValue.toString().isNotEmpty()) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "This should not exceed $cellValue characters",
                                            "localizableName" to "shouldNotExceed",
                                            "value" to cellValue,
                                        )
                                        serverValidations.add(validation)
                                    }
                                }
                                "minLength" -> {
                                    if (cellValue.toString().isNotEmpty()) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "This should not be less than $cellValue characters",
                                            "localizableName" to "shouldNotBeLess",
                                            "value" to cellValue,
                                        )
                                        serverValidations.add(validation)
                                    }
                                }
                                "pattern" -> {
                                    if (cellValue.toString().isNotEmpty()) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "The entered expression is not valid",
                                            "localizableName" to "expressionNotValid",
                                            "value" to cellValue,
                                        )
                                        serverValidations.add(validation)
                                    }
                                }
                                "email" -> {
                                    if (cellValue.equals("yes")) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "This field is invalid",
                                            "localizableName" to "fieldInvalid",
                                        )
                                        serverValidations.add(validation)
                                    }
                                }
                                "type" -> {
                                    if (cellValue.equals("select")) {
                                        rowData[columnName] = cellValue

                                        Log.d(
                                            "readExcelFile function",
                                            "***** drop down detected. Looking for its options...",
                                        )

                                        val optionsSheet =
                                            workbook.getSheet(localizableName + "_options")
                                        Log.d(
                                            "readExcelFile function",
                                            "***** Options Sheet Name : ${optionsSheet.sheetName}",
                                        )

                                        val options = mutableListOf<String>()

                                        for (x in 1..optionsSheet.lastRowNum) {
                                            val column = optionsSheet.getRow(x)
                                            val cellule = column.getCell(0)
                                            // Get the cell value according to its type
                                            val option = when (cellule.cellType) {
                                                CellType.NUMERIC -> formatter.formatCellValue(
                                                    cellule,
                                                )
                                                CellType.BOOLEAN -> cellule.booleanCellValue
                                                else -> cellule.stringCellValue
                                            }
                                            Log.d(
                                                "readExcelFile function",
                                                "***** option : $option",
                                            )
                                            options.add(option as String)
                                        }

                                        rowData["options"] = options
                                    } else {
                                        // if its not a dropdown we can simply add it with its corresponding value (input, datepicker...)
                                        rowData[columnName] = cellValue
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
        uri: Uri,
        xmlFiles: MutableList<Pair<Document, Element>>,
    ): MutableMap<String, Any> {
        // get the contentResolver
        val contentResolver = activity.applicationContext.contentResolver
        val inputStream = contentResolver.openInputStream(uri)

        // Create a POI File System object
        // val myFileSystem = POIFSFileSystem(inputStream)
        // Create a workbook using the File System
        // val workbook = HSSFWorkbook(myFileSystem)
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

                // if it is a sheet that contains only dropdown options we will not read it right now
                // because we read it only on-demand
                if (sheet.sheetName.contains("options")) {
                    continue
                }

                val fieldDetails = mutableMapOf<String, Any>()
                fieldDetails["label"] = sheet.sheetName
                Log.d("readExcelFile function", "+++++ sheet name : ${sheet.sheetName} +++++")

                // set the localizableName of this section (sheet)
                // remove spaces from the label of this section (sheet)
                var sectionLocalizableName = sheet.sheetName.replace("\\s".toRegex(), "")
                // remove 's from the label of this section (sheet)
                sectionLocalizableName = sectionLocalizableName.replace("'s", "")
                // Now we get the right localizableName, so we can add it
                fieldDetails["localizableName"] = sectionLocalizableName

                // Create a list to hold the rows
                val rowsList = mutableListOf<Map<String, Any>>()

                // Loop through the rows in the sheet
                for (i in 1 until sheet.physicalNumberOfRows) {
                    // Create a variable that will contain the localizableName of this field
                    var localizableName = ""

                    val row = sheet.getRow(i)
                    val rowData = mutableMapOf<String, Any>()
                    // create a list holding all the validations of the current line/element
                    val serverValidations = mutableListOf<Map<String, Any>>()

                    // Loop through the cells in the row
                    for (j in 0 until row.lastCellNum) {
                        val columnName = sheet.getRow(0).getCell(j).stringCellValue

                        // TODO : null cell not handled
                        if (columnName == null) {
                            Log.e(
                                "readExcelFile function",
                                "++++++ cell is null",
                            )
                        }

                        Log.d(
                            "readExcelFile function",
                            "+++++ column name : $columnName",
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
                            if (cellValue.equals("true")) {
                                cellValue = cell.stringCellValue.toBoolean()
                            }
                            if (cellValue.equals("false")) {
                                cellValue = cell.stringCellValue.toBoolean()
                            }

                            // Save the localizableName of this line
                            // so we can use it; if it requires options for the dropdown
                            if (columnName.equals("localizableName")) {
                                localizableName = cell.stringCellValue
                            }

                            // if this column is for internationalisation we should add its value to string.xml
                            // if this column is a validation one we should add its value to the validation list
                            // if the type of this field is select (dropdown) we should read the sheet that contains its options (using its localizableName)
                            when (columnName) {
                                "Français" -> {
                                    // get french file
                                    val (frenchXmlFile, frenchXmlFileRoot) = xmlFiles[0]
                                    // Add a string element
                                    val string = frenchXmlFile.createElement("string")
                                    string.setAttribute("name", localizableName)
                                    string.appendChild(frenchXmlFile.createTextNode(cellValue as String?))
                                    frenchXmlFileRoot.appendChild(string)
                                }
                                "English" -> {
                                    // get english file
                                    val (englishXmlFile, englishXmlFileRoot) = xmlFiles[1]
                                    // Add a string element
                                    val string = englishXmlFile.createElement("string")
                                    string.setAttribute("name", localizableName)
                                    string.appendChild(englishXmlFile.createTextNode(cellValue as String?))
                                    englishXmlFileRoot.appendChild(string)
                                }
                                "Portugais" -> {
                                    // get portuguese file
                                    val (portugueseXmlFile, portugueseXmlFileRoot) = xmlFiles[2]
                                    // Add a string element
                                    val string = portugueseXmlFile.createElement("string")
                                    string.setAttribute("name", localizableName)
                                    string.appendChild(portugueseXmlFile.createTextNode(cellValue as String?))
                                    portugueseXmlFileRoot.appendChild(string)
                                }
                                "required" -> {
                                    if (cellValue == true) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "This field is mandatory",
                                            "localizableName" to "fieldMandatory",
                                        )
                                        serverValidations.add(validation)
                                    } else {
                                        Log.d(
                                            "readExcelFile function",
                                            "/////////// Not Required ////////////",
                                        )
                                    }
                                }
                                "maxLength" -> {
                                    if (cellValue.toString().isNotEmpty()) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "This should not exceed $cellValue characters",
                                            "localizableName" to "shouldNotExceed",
                                            "value" to cellValue,
                                        )
                                        serverValidations.add(validation)
                                    } else {
                                        Log.d(
                                            "readExcelFile function",
                                            "/////////// No Max Length ////////////",
                                        )
                                    }
                                }
                                "minLength" -> {
                                    if (cellValue.toString().isNotEmpty()) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "This should not be less than $cellValue characters",
                                            "localizableName" to "shouldNotBeLess",
                                            "value" to cellValue,
                                        )
                                        serverValidations.add(validation)
                                    } else {
                                        Log.d(
                                            "readExcelFile function",
                                            "/////////// No Min Length ////////////",
                                        )
                                    }
                                }
                                "pattern" -> {
                                    if (cellValue.toString().isNotEmpty()) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "The entered expression is not valid",
                                            "localizableName" to "expressionNotValid",
                                            "value" to cellValue,
                                        )
                                        serverValidations.add(validation)
                                    } else {
                                        Log.d(
                                            "readExcelFile function",
                                            "/////////// No pattern ////////////",
                                        )
                                    }
                                }
                                "email" -> {
                                    if (cellValue.equals("yes")) {
                                        val validation = mutableMapOf(
                                            "name" to columnName,
                                            "message" to "This field is invalid",
                                            "localizableName" to "fieldInvalid",
                                        )
                                        serverValidations.add(validation)
                                    } else {
                                        Log.d(
                                            "readExcelFile function",
                                            "/////////// Not An Email ////////////",
                                        )
                                    }
                                }
                                "type" -> {
                                    if (cellValue.equals("select")) {
                                        rowData[columnName] = cellValue

                                        Log.d(
                                            "readExcelFile function",
                                            "***** drop down detected. Looking for its options...",
                                        )

                                        val optionsSheet =
                                            workbook.getSheet(localizableName + "_options")
                                        Log.d(
                                            "readExcelFile function",
                                            "***** Options Sheet Name : ${optionsSheet.sheetName}",
                                        )

                                        val options = mutableListOf<String>()

                                        for (x in 1..optionsSheet.lastRowNum) {
                                            val column = optionsSheet.getRow(x)
                                            val cellule = column.getCell(0)
                                            // Get the cell value according to its type
                                            val option = when (cellule.cellType) {
                                                CellType.NUMERIC -> formatter.formatCellValue(
                                                    cellule,
                                                )
                                                CellType.BOOLEAN -> cellule.booleanCellValue
                                                else -> cellule.stringCellValue
                                            }
                                            Log.d(
                                                "readExcelFile function",
                                                "***** option : $option",
                                            )
                                            options.add(option as String)
                                        }
                                        rowData["options"] = options
                                    } else {
                                        // if its not a dropdown we can simply add it with its corresponding value (input, datepicker...)
                                        rowData[columnName] = cellValue
                                    }
                                }
                                else -> {
                                    // Add the cell value to the row data
                                    rowData[columnName] = cellValue
                                }
                            }
                        }
                    }

                    rowData["serverValidations"] = serverValidations
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

        // Convert the Json object to a string
        val jsonString = gson.toJson(rowsList)

        // Print the string representation of the Json object
        Log.d("convertExcelToJson function", jsonString)

        return jsonString
    }

    private fun saveJsonFile(
        fileName: String = "demographics",
        json: String,
        activity: ComponentActivity,
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
