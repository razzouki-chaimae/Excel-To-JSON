package com.idemia.exceltojson

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.viewModels
import androidx.compose.foundation.layout.*
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.ArrowDropDown
import androidx.compose.material3.ButtonDefaults
import androidx.compose.material3.Icon
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.tooling.preview.Preview
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
                    modifier = Modifier.fillMaxSize(),
                )
            }
        }
        // val registry = this@MainActivity.activityResultRegistry
        // viewModel = ExcelToJsonViewModel(registry)
        viewModel.initialize(this)
        // activityResultRegistry.unregister("excel_file_picker")
    }

    @Composable
    fun ExcelToJsonScreen(modifier: Modifier = Modifier) {
        Column(
            modifier = modifier,
            verticalArrangement = Arrangement.Center,
            horizontalAlignment = Alignment.CenterHorizontally,
        ) {
            Upload()
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
                modifier = Modifier.size(ButtonDefaults.IconSize),
            )
            Spacer(modifier = Modifier.size(ButtonDefaults.IconSize))
            Text(text = "Upload excel file")
        }
    }

    private fun convertExcelIntoJson() {
        viewModel.chooseFile()
    }

    @Preview(showBackground = true)
    @Composable
    fun DefaultPreview() {
        ExcelToJsonTheme {
            ExcelToJsonScreen(Modifier.fillMaxSize())
        }
    }
}
