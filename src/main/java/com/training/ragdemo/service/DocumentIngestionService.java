package com.training.ragdemo.service;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.ai.document.Document;
import org.springframework.ai.reader.pdf.PagePdfDocumentReader;
import org.springframework.ai.transformer.splitter.TokenTextSplitter;
import org.springframework.ai.vectorstore.VectorStore;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * Reads the bundled HR policy PDF, splits it into chunks, and stores the chunk
 * embeddings in the MariaDB vector store. This is the "Document Ingestion &
 * Processing" + "Embeddings & Vector Storage" step of the RAG pipeline.
 */
@Service
public class DocumentIngestionService {

    private static final Logger log = LoggerFactory.getLogger(DocumentIngestionService.class);

    private final VectorStore vectorStore;
    private final JdbcTemplate jdbcTemplate;
    private final Resource[] documentResources;

    public DocumentIngestionService(VectorStore vectorStore,
                                     JdbcTemplate jdbcTemplate,
                                     @Value("${rag.document-paths}") Resource[] documentResources) {
        this.vectorStore = vectorStore;
        this.jdbcTemplate = jdbcTemplate;
        this.documentResources = documentResources;
    }

    public int ingest() {
        int total = 0;
        for (Resource resource : documentResources) {
            log.info("Reading PDF: {}", resource.getFilename());
            PagePdfDocumentReader pdfReader = new PagePdfDocumentReader(resource);
            List<Document> pages = pdfReader.get();
            pages.forEach(page -> page.getMetadata().put("source", resource.getFilename()));

            List<Document> chunks = new TokenTextSplitter().apply(pages);
            log.info("Split {} page(s) of {} into {} chunk(s)", pages.size(), resource.getFilename(), chunks.size());

            vectorStore.add(chunks);
            total += chunks.size();
        }
        log.info("Indexed {} chunk(s) total into the vector store", total);
        return total;
    }

    public long countStoredChunks() {
        Long count = jdbcTemplate.queryForObject("SELECT COUNT(*) FROM vector_store", Long.class);
        return count == null ? 0L : count;
    }
}
