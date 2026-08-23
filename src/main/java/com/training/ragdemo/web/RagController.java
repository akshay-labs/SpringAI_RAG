package com.training.ragdemo.web;

import com.training.ragdemo.service.DocumentIngestionService;
import com.training.ragdemo.service.RagQueryService;
import com.training.ragdemo.web.dto.CompareResponse;
import com.training.ragdemo.web.dto.IngestResponse;
import com.training.ragdemo.web.dto.QueryRequest;
import com.training.ragdemo.web.dto.QueryResponse;
import com.training.ragdemo.web.dto.StatusResponse;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/v1")
public class RagController {

    private final DocumentIngestionService ingestionService;
    private final RagQueryService ragQueryService;

    public RagController(DocumentIngestionService ingestionService, RagQueryService ragQueryService) {
        this.ingestionService = ingestionService;
        this.ragQueryService = ragQueryService;
    }

    @PostMapping("/documents/ingest")
    public IngestResponse ingest() {
        return new IngestResponse(ingestionService.ingest());
    }

    @GetMapping("/documents/status")
    public StatusResponse status() {
        return new StatusResponse(ingestionService.countStoredChunks());
    }

    @PostMapping("/rag/query")
    public QueryResponse query(@RequestBody QueryRequest request) {
        RagQueryService.Result result = ragQueryService.answerWithSources(request.question());
        return new QueryResponse(result.answer(), result.sources());
    }

    @PostMapping("/rag/compare")
    public CompareResponse compare(@RequestBody QueryRequest request) {
        String directAnswer = ragQueryService.answerDirect(request.question());
        RagQueryService.Result ragResult = ragQueryService.answerWithSources(request.question());
        return new CompareResponse(request.question(), directAnswer, ragResult.answer(), ragResult.sources());
    }
}
