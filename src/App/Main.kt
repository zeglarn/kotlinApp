package App

import com.jfoenix.controls.JFXButton
import javafx.application.Application
import javafx.geometry.Insets
import javafx.scene.Scene
import javafx.scene.layout.VBox
import javafx.stage.FileChooser
import javafx.stage.Stage

import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.awt.Desktop
import java.io.File
import java.io.FileInputStream


class Main : Application() {
    override fun start(primaryStage: Stage) = try {
        val root = VBox()
        root.padding = Insets(5.0,5.0,5.0,5.0)
        //val desktop = Desktop.getDesktop()
        val filechooser = FileChooser()

        val openButton = JFXButton("Click to select file.")

        openButton.setOnAction { _ ->
            val file = filechooser.showOpenDialog(primaryStage)
            if (file != null) println(file.absolutePath)


            val vb = XSSFWorkbook(file)
            val sheet = vb.getSheetAt(0)

            val rows = sheet.getPhysicalNumberOfRows()
            var cols = 0

            for (i in 0..rows) {
                val row = sheet.getRow(i)
                if (row != null) {
                    var tmp = sheet.getRow(i).getPhysicalNumberOfCells()
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
                            print(cell.rawValue)
                            print("\t")
                        }
                    }
                }
                println("")
            }


        }

        root.children.addAll(openButton)


        val scene = Scene(root, 200.0, 50.0)
        primaryStage.scene = scene
        primaryStage.show()
    } catch (e: Exception) { e.printStackTrace() }
}

fun main(args: Array<String>) {
    Application.launch(Main::class.java, *args)
}