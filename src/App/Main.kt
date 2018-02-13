package App

import com.itextpdf.text.*
import com.itextpdf.text.pdf.PdfWriter
import com.jfoenix.controls.JFXButton
import javafx.application.Application
import javafx.geometry.Insets
import javafx.scene.Scene
import javafx.scene.layout.VBox
import javafx.stage.FileChooser
import javafx.stage.Stage
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream


class Main : Application() {
    override fun start(primaryStage: Stage) = try {
        val root = VBox()
        root.padding = Insets(5.0,5.0,5.0,5.0)
        val fileChooser = FileChooser()

        val openButton = JFXButton("Click to select file.")

        openButton.setOnAction { _ ->
            val file = fileChooser.showOpenDialog(primaryStage)
            if (file != null) println(file.absolutePath)
            val result = StringBuilder()

            val vb = XSSFWorkbook(file)
            val sheet = vb.getSheetAt(0)

            val rows = sheet.physicalNumberOfRows
            var cols = 0

            for (i in 0..rows) {
                val row = sheet.getRow(i)
                if (row != null) {
                    val tmp = sheet.getRow(i).physicalNumberOfCells
                    if (tmp > cols) {
                        cols = tmp
                    }
                }
            }

            for (r in 0..rows) {
                val row = sheet.getRow(r)
                if (row != null) {
                    for (c in 0..cols) {
                        val cell = row.getCell(c)
                        if (cell != null) {
                            result.append(cell.rawValue)
                            result.append("   ")
                            print(cell.rawValue)
                            print("\t")
                        }
                    }
                }
                result.append("\n")
                println("")
            }
            createPDF("./test.pdf", result.toString())
        }
        root.children.addAll(openButton)

        val scene = Scene(root, 200.0, 50.0)
        primaryStage.scene = scene
        primaryStage.show()
    } catch (e: Exception) { e.printStackTrace() }

    private fun createPDF(dest: String, content: String) {
        val document = Document(PageSize.A4)
        val writer = PdfWriter.getInstance(document, FileOutputStream(dest))
        writer.setPdfVersion(PdfWriter.PDF_VERSION_1_7)
        document.open()
        val paragraph = Paragraph()
        paragraph.font = Font(Font.FontFamily.HELVETICA, 20f)
        val chunk = Chunk(content)
        paragraph.add(chunk)
        document.add(paragraph)
        document.close()
    }
}

fun main(args: Array<String>) {
    Application.launch(Main::class.java, *args)
}