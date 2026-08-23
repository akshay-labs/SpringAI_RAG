package com.training.ragdemo.service;

import org.springframework.ai.chat.client.ChatClient;
import org.springframework.ai.document.Document;
import org.springframework.ai.vectorstore.SearchRequest;
import org.springframework.ai.vectorstore.VectorStore;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.util.List;

import com.training.ragdemo.web.dto.SourceChunk;

/**
 * The "Retrieval & Context Generation" + "LLM Integration" steps of the RAG
 * pipeline, done explicitly (rather than via QuestionAnswerAdvisor) so the
 * retrieval step is visible and the matched chunks can be returned to the caller.
 */
@Service
public class RagQueryService {

    private static final String SYSTEM_PROMPT = """
            You are an HR assistant for Northwind Retail Group. Answer the user's question using ONLY
            the HR policy context provided below. If the answer is not contained in the context, say
            you don't know and suggest the employee contact HR - do not make anything up.

            Context:
            %s
            """;

    private final ChatClient chatClient;
    private final VectorStore vectorStore;
    private final int topK;

    public RagQueryService(ChatClient.Builder chatClientBuilder,
                            VectorStore vectorStore,
                            @Value("${rag.top-k}") int topK) {
        this.chatClient = chatClientBuilder.build();
        this.vectorStore = vectorStore;
        this.topK = topK;
    }

    public Result answerWithSources(String question) {
        List<Document> matches = vectorStore.similaritySearch(
                SearchRequest.builder().query(question).topK(topK).build());

        String context = matches.isEmpty()
                ? "(no matching HR policy content found)"
                : matches.stream().map(Document::getText).reduce("", (a, b) -> a + "\n---\n" + b);

        String answer = chatClient.prompt()
                .system(SYSTEM_PROMPT.formatted(context))
                .user(question)
                .call()
                .content();

        List<SourceChunk> sources = matches.stream()
                .map(doc -> new SourceChunk(
                        String.valueOf(doc.getMetadata().getOrDefault("source", "unknown")),
                        excerpt(doc.getText()),
                        doc.getScore() == null ? 0.0 : doc.getScore()))
                .toList();

        return new Result(answer, sources);
    }

    /**
     * Calls the LLM directly with no retrieval and no HR context, for the "RAG vs. no-RAG"
     * comparison demo - this is what the model says when it has to rely on its own training data.
     */
    public String answerDirect(String question) {
        return chatClient.prompt()
                .user(question)
                .call()
                .content();
    }

    private String excerpt(String text) {
        String trimmed = text.strip();
        return trimmed.length() > 220 ? trimmed.substring(0, 220) + "..." : trimmed;
    }

    public record Result(String answer, List<SourceChunk> sources) {
    }
}
