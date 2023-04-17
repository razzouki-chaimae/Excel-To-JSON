package com.idemia.exceltojson

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.viewModels
import androidx.compose.foundation.layout.*
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.ArrowDropDown
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.runtime.livedata.observeAsState
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.tooling.preview.Preview
import androidx.compose.ui.unit.dp
import com.idemia.exceltojson.ui.theme.ExcelToJsonTheme
import com.idemia.exceltojson.viewModel.ExcelToJsonViewModel

class MainActivity : ComponentActivity() {

    // get the ViewModel instance
    private val viewModel by viewModels<ExcelToJsonViewModel>()
    
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent {
            ExcelToJsonTheme {
                ExcelToJsonScreen(
                    modifier = Modifier.fillMaxSize()
                )
            }
        }
        viewModel.initialize(this)
    }

    @Composable
    fun ExcelToJsonScreen(modifier: Modifier = Modifier) {

        Column(
            modifier = modifier,
            verticalArrangement = Arrangement.Center,
            horizontalAlignment = Alignment.CenterHorizontally
        ) {
            Upload()
            OptionsDialog(viewModel = viewModel)
        }
    }

    @Composable
    fun Upload() {
        OutlinedButton(onClick = {
            convertExcelIntoJson()
        }) {
            Icon(
                imageVector = Icons.Outlined.ArrowDropDown,
                contentDescription = "Upload Icon",
                modifier = Modifier.size(ButtonDefaults.IconSize)
            )
            Spacer(modifier = Modifier.size(ButtonDefaults.IconSize))
            Text(text = "Upload excel file")
        }
    }

    private fun convertExcelIntoJson() {
        viewModel.chooseFile()
    }

    @Composable
    fun OptionsDialog(viewModel: ExcelToJsonViewModel) {
        //val showDialog = remember { mutableStateOf(false) }
        //val showDialog by viewModel.showDialog.collectAsState()
        val showDialog by viewModel.showDialog.observeAsState(initial = false)
        val message by viewModel.dialogMessage.observeAsState(initial = "")

        if (showDialog) {
            AlertDialog(
                onDismissRequest = {
                    //viewModel.dismissDialog()
                    viewModel.onDialogDismissed()
                                   },
                title = { Text(text = "Select options") },
                text = { Text(text = "Please import $message's options.") },
                confirmButton = {
                    Button(
                        onClick = {
                            // Do something when the button is clicked
                            //viewModel.confirmDialog()
                            viewModel.onDialogConfirmed()
                        },
                        modifier = Modifier
                            .padding(vertical = 8.dp)
                            .fillMaxWidth()
                    ) {
                        Text(text = "Import")
                    }
                }
            )
        }
    }

    fun showDialog(viewModel: ExcelToJsonViewModel){
        //viewModel.showDialog()
    }

    /*@Composable
    fun ExampleDialog(onDismiss: () -> Unit) {
        var showDialog by remember { mutableStateOf(true) }
        if (showDialog) {
            AlertDialog(
                onDismissRequest = onDismiss,
                title = { Text("Example Dialog") },
                text = { Text("This is an example dialog.") },
                buttons = {
                    Row(
                        modifier = Modifier.padding(all = 8.dp),
                        horizontalArrangement = Arrangement.End
                    ) {
                        Button(onClick = { showDialog = false }) {
                            Text("OK")
                        }
                    }
                }
            )
        }
    }

    @Composable
    fun optionAlert(viewModel: ExcelToJsonViewModel, cell : String) {
        val showDialog = remember { mutableStateOf(false) }

        if (showDialog.value) {
            AlertDialog(
                onDismissRequest = { showDialog.value = false },
                title = { Text("Specify options") },
                text = { Text("Please, give me the $cell's options.") },
                confirmButton = {
                    Button(onClick = { showDialog.value = false }) {
                        Text("OK")
                    }
                }
            )
        }

        // ... rest of the composable code

        Button(
            onClick = { viewModel.showDialog(showDialog) }
        ) {
            Text("Show Dialog")
        }
    }*/

    @Preview(showBackground = true)
    @Composable
    fun DefaultPreview() {
        ExcelToJsonTheme {
            ExcelToJsonScreen(Modifier.fillMaxSize())
        }
    }
}