package org.dba.xdtest

import javafx.application.Platform
import javafx.beans.property.SimpleStringProperty
import javafx.collections.FXCollections
import javafx.fxml.FXML
import javafx.scene.control.Button
import javafx.scene.control.Label
import javafx.scene.control.RadioButton
import javafx.scene.control.TableCell
import javafx.scene.control.TableColumn
import javafx.scene.control.TableView
import javafx.scene.control.ToggleGroup
import javafx.scene.text.Font
import javafx.stage.FileChooser
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.SupervisorJob
import kotlinx.coroutines.javafx.JavaFx
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.slf4j.Logger
import org.slf4j.LoggerFactory
import java.io.File
import java.io.IOException
import java.util.prefs.Preferences

private val rackConfig = FileParseConfig(
    sheetName = "XD 700",
    firstRow = 4,
    lastRow = 31,
    itemCol = 6,
    qtyCol = 7,
    skuCol = 8,
    solIDCol = 10
)

private val CDUConfig = FileParseConfig(
    sheetName = "XD 700",
    firstRow = 49,
    lastRow = 55,
    itemCol = 6,
    qtyCol = 7,
    skuCol = 8,
    solIDCol = 10
)

private val partsConfig = FileParseConfig(
    sheetName = "XD 700",
    firstRow = 42,
    lastRow = 47,
    itemCol = 6,
    qtyCol = 7,
    skuCol = 8,
    solIDCol = 10
)

private val targetFileConfig = FileParseConfig(
    sheetName = "ExpertBOM",
    firstRow = 6,
    lastRow = 0,
    itemCol = 0,
    qtyCol = 1,
    skuCol = 2,
    solIDCol = 5
)

private val prefs: Preferences = Preferences.userNodeForPackage(MainController::class.java)
private const val lastTargetDirKey = "lastTargetDir"
private lateinit var targetFile: File

private val controllerScope = CoroutineScope(Dispatchers.JavaFx + SupervisorJob())
private val dataFormatter = DataFormatter()
private val logger: Logger = LoggerFactory.getLogger("XD test")

lateinit var sourceConfig: FileParseConfig

private val configFilePath = "C:\\Users\\albertd\\OneDrive - Hewlett Packard Enterprise\\HPE\\Cray XD\\XD700 Config Builder v2.0.xlsx"
private var sourceFile = File(configFilePath)

private val diffList = mutableListOf<DiffResult>()

class MainController {
    lateinit var buttonTarget: Button

    @FXML
    lateinit var targetSolIDCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var sourceSolIDCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var targetQtyCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var sourceQtyCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var targetItemCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var sourceItemCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var skuCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var tableDiff: TableView<DiffResult>

    @FXML
    lateinit var labelStatus: Label

    @FXML
    lateinit var radioParts: RadioButton

    @FXML
    lateinit var radioCDU: RadioButton

    @FXML
    lateinit var confGroup: ToggleGroup

    @FXML
    lateinit var radioRack: RadioButton

    @FXML
    lateinit var buttonQuit: Button

    @FXML
    lateinit var buttonCompare: Button

    @FXML
    lateinit var labelTargetFile: Label

    @FXML
    lateinit var labelSourceFile: Label

    @FXML
    lateinit var labelJavaFX: Label

    @FXML
    lateinit var labelJDK: Label

    fun initialize() {
        labelJDK.text = "Java SDK version: ${Runtime.version()}"
        labelJavaFX.text = "JavaFX version: ${System.getProperty("javafx.runtime.version")}"

        tableDiff.columns.setAll(
            makeColumn("SKU", 200.0, { it.sku }),
            makeColumn("Source Qty", 100.0, { it.sourceQty }) { it.sourceQty != it.targetQty },
            makeColumn("Target Qty", 100.0, { it.targetQty }) { it.sourceQty != it.targetQty },
            makeColumn("Source Item", 100.0, { it.sourceItem }) { it.sourceItem != it.targetItem },
            makeColumn("Target Item", 100.0, { it.targetItem }) { it.sourceItem != it.targetItem },
            makeColumn("Source SolID", 100.0, { it.sourceSolID }) { it.sourceSolID != it.targetSolID },
            makeColumn("Target SolID", 100.0, { it.targetSolID }) { it.sourceSolID != it.targetSolID }
        )
        tableDiff.isVisible = false
        sourceConfig = rackConfig
        labelSourceFile.text = "Rack"
    }

    private fun <T> makeColumn(
        title: String,
        width: Double,
        extractor: (DiffResult) -> T?,
        highlight: (DiffResult) -> Boolean = { false }
    ): TableColumn<DiffResult, String> {

        val col = TableColumn<DiffResult, String>(title)
        col.setCellValueFactory { cell ->
            SimpleStringProperty(extractor(cell.value)?.toString() ?: "")
        }

        col.prefWidth = width
        col.setCellFactory {
            object : TableCell<DiffResult, String>() {
                override fun updateItem(item: String?, empty: Boolean) {
                    super.updateItem(item, empty)

                    styleClass.removeAll("diff-cell", "match-cell")

                    if (empty) {
                        text = ""
                        return
                    }
                    val row = tableView.items[index]
                    text = item

                    if (highlight(row)) {
                        styleClass.add("diff-cell")
                    } else {
                        styleClass.add("match-cell")
                    }
                }
            }
        }
        return col
    }
    @FXML
    fun sourceSelect() {
        var selectedSource = ""
        val selectedToggle = confGroup.selectedToggle
        if (selectedToggle != null) {
            val selectedRadio = selectedToggle as RadioButton
            selectedSource = selectedRadio.text
        }

        when (selectedSource) {
            "Rack" -> {
                sourceConfig = rackConfig
            }

            "CDU" -> {
                sourceConfig = CDUConfig
            }

            "Parts" -> {
                sourceConfig = partsConfig
            }
        }
        labelSourceFile.text = selectedSource
    }

    private fun readValuesFromFile(file: File, config: FileParseConfig): List<ValueList> {
        val values = mutableListOf<ValueList>()
        try {
            WorkbookFactory.create(file, null, true).use { workbook ->
                val evaluator = workbook.creationHelper.createFormulaEvaluator()
                val sheet = workbook.getSheet(config.sheetName)
                val lastRow = if (config.lastRow == 0) sheet.lastRowNum else config.lastRow


                for (i in config.firstRow..lastRow) {
                    val row = sheet.getRow(i) ?: continue

                    val itemValue = dataFormatter.formatCellValue(row.getCell(config.itemCol), evaluator).trim()
                    if (itemValue.isBlank() || itemValue.equals("Total", ignoreCase = true)) continue

                    val qtyValue = dataFormatter.formatCellValue(row.getCell(config.qtyCol), evaluator).trim()
                    val skuValue = dataFormatter.formatCellValue(row.getCell(config.skuCol), evaluator).trim()
                    val solIDValue = dataFormatter.formatCellValue(row.getCell(config.solIDCol), evaluator).trim()

                    values.add(ValueList(itemValue, qtyValue, skuValue, solIDValue))
                }
                workbook.close()
            }
        } catch (e: IOException) {
            logger.error("Error reading file ${file.name}", e)
            throw e
        }

        return values
    }

    @FXML
    fun handleCompare() {
        tableDiff.isVisible = false
        diffList.clear()
        controllerScope.launch {
            try {
                withContext(Dispatchers.JavaFx) {
                    val configValues = readValuesFromFile(sourceFile, sourceConfig)
                    labelStatus.text = "Source File Read"
                    logger.info("File $sourceFile Read ")

                    val targetValues = readValuesFromFile(targetFile, targetFileConfig)


                    logger.info("File $targetFile Read ")

                    val mapTarget = targetValues.filter { it.sku != null }.associateBy { it.sku }
                    val mapConfig = configValues.filter { it.sku != null }.associateBy { it.sku }

                    for ((sku, source) in mapConfig) {
                        val target = mapTarget[sku]
                        if (target == null) {
                            diffList += DiffResult(
                                source.sku,source.item, null,source.quantity, null,source.solID, null
                            )
                        } else if (source != target) {
                            diffList += DiffResult(
                                source.sku,source.item, target.item,source.quantity, target.quantity,source.solID, target.solID
                            )
                        }
                    }

                    for ((sku, target) in mapTarget) {
                        if (sku !in mapConfig) {
                            diffList += DiffResult(
                                target.sku,
                                null, target.item, null, target.quantity, null, target.solID

                            )
                        }
                    }

                   logger.info("Compare Complete")

                    if (diffList.isNotEmpty()) {
                        tableDiff.isVisible = true
                        tableDiff.items = FXCollections.observableArrayList(diffList)
                        labelStatus.text = "File Mismatch"
                        labelStatus.font = Font.font(24.0)
                        labelStatus.textFill = javafx.scene.paint.Color.RED

                    } else {
                        labelStatus.text = "File Match"
                        labelStatus.font = Font.font(24.0)
                        labelStatus.textFill = javafx.scene.paint.Color.GREEN
                    }
                }

            }
            catch (e: Exception) {
                logger.error("A critical error occurred during file processing.", e)
                labelStatus.text = "ERROR: An unexpected error occurred. Check logs"
            }
        }

    }

    @FXML
    fun handleOpenTargetFile() {
        val initialDir = prefs.get(lastTargetDirKey, System.getProperty("user.home"))
        openFileChooser(initialDir)?.let { file ->
            targetFile = file
            labelTargetFile.text = file.name
            prefs.put(lastTargetDirKey, file.parent) // Save the new directory
        }
    }

    private fun openFileChooser(directoryPath: String): File? {
        val fileChooser = FileChooser().apply {
            title = "Open Target Excel File"
            initialDirectory = File(directoryPath)
            extensionFilters.add(FileChooser.ExtensionFilter("Excel Files (*.xlsx)", "*.xlsx"))
        }
        return fileChooser.showOpenDialog(labelJavaFX.scene.window)
    }

    @FXML
    fun handleQuit() {
        Platform.exit()

    }

}