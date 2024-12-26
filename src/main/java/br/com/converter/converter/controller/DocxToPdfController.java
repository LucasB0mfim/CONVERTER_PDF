package br.com.converter.converter.controller;

import java.io.IOException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import br.com.converter.converter.service.DocxToPdfService;

/**
 * @author Lucas Bomfim 
 */

@RestController
@RequestMapping("/convert")
@CrossOrigin(origins = "*")
public class DocxToPdfController {

    @Autowired
    private DocxToPdfService docxToPdfService;

    @PostMapping("/docx-to-pdf")
    public ResponseEntity<byte[]> convertDocxToPdf(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return ResponseEntity.badRequest().body("Arquivo vazio.".getBytes());
        }

        try {
            byte[] pdfBytes = docxToPdfService.convertDocxToPdfBytes(file.getInputStream());

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_PDF);
            headers.setContentDisposition(ContentDisposition.builder("attachment")
                    .filename(file.getOriginalFilename().replace(".docx", ".pdf"))
                    .build());

            return ResponseEntity.ok().headers(headers).body(pdfBytes);

        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.internalServerError().body("Erro ao converter o arquivo.".getBytes());
        }
    }
}
