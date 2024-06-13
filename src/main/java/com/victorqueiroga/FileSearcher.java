package com.victorqueiroga;

import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Scanner;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class FileSearcher {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        System.out.println("Digite o caminho do diretório a ser vasculhado:");
        String directoryPath = scanner.nextLine();
        System.out.println("Digite o valor a ser procurado:");
        String searchValue = scanner.nextLine().toLowerCase();
        scanner.close();

        try {
            Files.walkFileTree(Paths.get(directoryPath), new SimpleFileVisitor<Path>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    String fileName = file.toString().toLowerCase();
                    
                        if (fileName.endsWith(".doc") || fileName.endsWith(".docx")) {
                            if (containsWord(file, searchValue)) {
                                System.out.println("Valor encontrado em: " + file);
                            }
                        } else if (fileName.endsWith(".xls") || fileName.endsWith(".xlsx")) {
                            if (containsExcel(file, searchValue)) {
                                System.out.println("Valor encontrado em: " + file);
                            }
                        } else if (fileName.endsWith(".pdf")) {
                            if (containsPDF(file, searchValue)) {
                                System.out.println("Valor encontrado em: " + file);
                            }
                        }
                  
                    return FileVisitResult.CONTINUE;
                }
            });
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean containsWord(Path file, String searchValue) {
        try {
            if (file.toString().endsWith(".doc")) {
                try (HWPFDocument doc = new HWPFDocument(Files.newInputStream(file));
                     WordExtractor extractor = new WordExtractor(doc)) {
                    String text = extractor.getText().toLowerCase();
                    return text.contains(searchValue);
                }
            } else if (file.toString().endsWith(".docx")) {
                try (XWPFDocument docx = new XWPFDocument(Files.newInputStream(file))) {
                    for (XWPFParagraph paragraph : docx.getParagraphs()) {
                        String text = paragraph.getText().toLowerCase();
                        if (text.contains(searchValue)) {
                            return true;
                        }
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("Erro ao processar o arquivo Word: " + file + " - " + e.getMessage());
        }
        return false;
    }

    private static boolean containsExcel(Path file, String searchValue) {
        try {
            if (file.toString().endsWith(".xls")) {
                try (HSSFWorkbook workbook = new HSSFWorkbook(Files.newInputStream(file))) {
                    return searchWorkbook(workbook, searchValue);
                }
            } else if (file.toString().endsWith(".xlsx")) {
                try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(file))) {
                    return searchWorkbook(workbook, searchValue);
                }
            }
        } catch (IOException e) {
            System.err.println("Erro ao processar o arquivo Excel: " + file + " - " + e.getMessage());
        }
        return false;
    }

    private static boolean searchWorkbook(Workbook workbook, String searchValue) {
        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().toLowerCase().contains(searchValue)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private static boolean containsPDF(Path file, String searchValue) {
        try (PDDocument document = PDDocument.load(Files.newInputStream(file))) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document).toLowerCase();
            return text.contains(searchValue);
        } catch (IOException e) {
            System.err.println("Erro ao processar o arquivo PDF: " + file + " - " + e.getMessage());
            return false;
        }
    }
}
