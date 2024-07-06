package com.example

import java.io.File
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.Properties
import java.util.logging.Level
import java.util.logging.Logger
import javax.mail.Session
import javax.mail.internet.MimeMessage
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Row

fun main() {
    val logger = Logger.getLogger("EMLToExcel")

    try {
        val emlFiles = chooseEmlFiles(logger)

        if (emlFiles.isNotEmpty()) {
            val extractedData = emlFiles.mapNotNull { parseEmlFile(it, logger) }.associateBy { it["Email"]!! }

            writeToExcel(extractedData.values.toList(), emlFiles[0].parentFile, logger)
        } else {
            logger.info("Nema pronađenih EML datoteka.")
        }
    } catch (e: Exception) {
        logger.log(Level.SEVERE, "Dogodila se pogreška u glavnom procesu", e)
    }
}

fun chooseEmlFiles(logger: Logger): List<File> {
    return try {
        val dir = File(System.getProperty("user.dir"))
        val files = dir.listFiles { _, name -> name.endsWith(".eml") }?.toList() ?: emptyList()
        logger.info("Pronađeno ${files.size} EML datoteka.")
        files
    } catch (e: Exception) {
        logger.log(Level.SEVERE, "Nije uspjelo odabiranje EML datoteka", e)
        emptyList()
    }
}

fun parseEmlFile(file: File, logger: Logger): Map<String, String>? {
    return try {
        val properties = Properties()
        val session = Session.getDefaultInstance(properties, null)
        val message = MimeMessage(session, file.inputStream())

        val dateFormat = SimpleDateFormat("dd.MM.yyyy HH:mm:ss")
        val date = message.sentDate?.let { dateFormat.format(it) } ?: ""
        val from = message.from?.joinToString(", ") { it.toString() } ?: ""

        val content = message.content.toString()
        val name = Regex("(?<=<td style=\"color:#555555;padding-top: 3px;padding-bottom: 20px;\">)(.*?)(?=</td>)").find(content)?.value ?: ""
        val email = Regex("(?<=<a href=\"mailto:)(.*?)(?=\">)").find(content)?.value ?: ""
        val phone = Regex("(?<=<td style=\"color:#555555;padding-top: 3px;padding-bottom: 20px;\">)(.*?)(?=</td>)").findAll(content).elementAtOrNull(2)?.value ?: ""

        if (name.isNotEmpty() && email.isNotEmpty() && phone.isNotEmpty()) {
            mapOf("Datum" to date, "Ime" to name, "Broj Mobitela" to phone, "Email" to email)
        } else {
            logger.warning("Ime, broj mobitela ili email nisu pronađeni u datoteci: ${file.name}")
            null
        }
    } catch (e: Exception) {
        logger.log(Level.SEVERE, "Nije uspjelo parsiranje EML datoteke: ${file.name}", e)
        null
    }
}

fun writeToExcel(data: List<Map<String, String>>, outputDir: File, logger: Logger) {
    try {
        val workbook: Workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Informacije Polaznika")

        val headerRow = sheet.createRow(0)
        val headers = listOf("Datum", "Ime", "Broj Mobitela", "Email")
        for ((i, header) in headers.withIndex()) {
            val cell = headerRow.createCell(i)
            cell.setCellValue(header)
        }

        for ((i, rowData) in data.withIndex()) {
            val row: Row = sheet.createRow(i + 1)
            for ((j, key) in headers.withIndex()) {
                val cell = row.createCell(j)
                cell.setCellValue(rowData[key])
            }
        }

        val outputFile = File(outputDir, "Informacije Polaznika.xlsx")
        val fileOut = FileOutputStream(outputFile)
        workbook.write(fileOut)
        fileOut.close()
        workbook.close()
        logger.info("Uspješno zapisani podaci u ${outputFile.absolutePath}")
    } catch (e: Exception) {
        logger.log(Level.SEVERE, "Nije uspjelo pisanje u Excel datoteku", e)
    }
}