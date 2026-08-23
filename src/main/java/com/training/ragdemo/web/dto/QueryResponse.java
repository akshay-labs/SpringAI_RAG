package com.training.ragdemo.web.dto;

import java.util.List;

public record QueryResponse(String answer, List<SourceChunk> sources) {
}
