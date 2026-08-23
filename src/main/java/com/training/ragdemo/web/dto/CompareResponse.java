package com.training.ragdemo.web.dto;

import java.util.List;

public record CompareResponse(String question, String directAnswer, String ragAnswer, List<SourceChunk> sources) {
}
