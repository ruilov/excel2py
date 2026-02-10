"""Prompt templates for AI-assisted refactor stage.

This module is intentionally plain text + small render helpers so prompts are:
- versioned in git
- easy to diff/review
- reusable across providers/models
"""

from __future__ import annotations

from textwrap import dedent


REFRACTOR_SYSTEM_PROMPT = dedent(
    """
    You are assisting with spreadsheet-to-Python refactoring.

    Hard constraints:
    - Preserve semantics exactly.
    - Do not invent business logic.
    - Prefer refactoring repeated formula patterns into clear functions/loops.
    - Keep traceability from refactored code back to source workbook cells/ranges.
    - Output only the requested schema/format, no extra prose.

    Acceptance rule:
    - Refactor is accepted only if parity/circular/determinism tests pass.
    """
).strip()


GLOBAL_CONTEXT_PROMPT_TEMPLATE = dedent(
    """
    Task:
    Summarize workbook intent and major calculation blocks for refactoring.

    Input JSON:
    {global_context_json}

    Produce JSON with this shape:
    {{
      "workbook_summary": "short summary of model purpose",
      "sheet_summaries": [
        {{"sheet": "name", "role": "what this sheet does"}}
      ],
      "major_blocks": [
        {{
          "block_id": "stable id",
          "sheet": "sheet name",
          "range_hint": "A1:D50 style",
          "label_evidence": ["label1", "label2"],
          "purpose_hypothesis": "what this block likely computes"
        }}
      ],
      "naming_glossary": [
        {{"raw_label": "original label", "suggested_name": "snake_case_name"}}
      ],
      "uncertainties": ["specific ambiguity 1", "specific ambiguity 2"]
    }}

    Rules:
    - Keep hypotheses conservative.
    - If unsure, state uncertainty explicitly.
    - Do not reference cells not present in input.
    """
).strip()


CLUSTER_PACKET_PROMPT_TEMPLATE = dedent(
    """
    Task:
    Propose reusable abstractions for repeated formula clusters.

    Global summary JSON:
    {global_summary_json}

    Cluster packet JSON:
    {cluster_packet_json}

    Produce JSON with this shape:
    {{
      "cluster_id": "from input",
      "proposed_function_name": "snake_case",
      "proposed_signature": ["arg1", "arg2", "arg3"],
      "abstraction_strategy": "loop|vector-like loop|helper extraction|keep explicit",
      "generalized_rule": "plain language rule",
      "lineage_requirements": ["what lineage comments/maps must be preserved"],
      "risk_flags": ["possible semantic risk 1"],
      "confidence": 0.0
    }}

    Rules:
    - Prefer simple abstractions over clever ones.
    - If cluster should stay explicit, say so explicitly.
    - Do not change formula semantics.
    """
).strip()


PLAN_SYNTHESIS_PROMPT_TEMPLATE = dedent(
    """
    Task:
    Merge cluster-level proposals into one executable refactor plan.

    Global summary JSON:
    {global_summary_json}

    Cluster proposals JSON:
    {cluster_proposals_json}

    Produce JSON with this shape:
    {{
      "functions": [
        {{
          "function_name": "snake_case",
          "covers_clusters": ["cluster_1", "cluster_2"],
          "inputs": ["arg_a", "arg_b"],
          "outputs": ["target cells/groups"],
          "implementation_style": "explicit|loop",
          "lineage_plan": "how source cell mapping is preserved"
        }}
      ],
      "orchestration_order": ["function_a", "function_b"],
      "clusters_left_explicit": ["cluster_x"],
      "open_questions": ["question 1"],
      "high_risk_areas": ["risk 1"]
    }}

    Rules:
    - Every cluster must be either covered or explicitly left explicit.
    - Keep plan deterministic and test-gated.
    - Avoid deep nesting and unnecessary indirection.
    """
).strip()


CODEGEN_PROMPT_TEMPLATE = dedent(
    """
    Task:
    Generate refactored Python code from deterministic literal source + approved plan.

    Literal source code:
    {literal_source_code}

    Approved refactor plan JSON:
    {approved_plan_json}

    Output:
    - Python code only.

    Rules:
    - Preserve runtime behavior.
    - Keep/emit lineage comments at function or assignment level.
    - Keep public entrypoint compatible with existing run_model API.
    - Do not add external dependencies.
    - Keep code readable and compact.
    """
).strip()


SELF_CHECK_PROMPT_TEMPLATE = dedent(
    """
    Task:
    Review proposed refactor for semantic risks before tests run.

    Refactored code:
    {refactored_code}

    Return JSON:
    {{
      "semantic_risk_findings": [
        {{"severity": "high|medium|low", "finding": "text", "location_hint": "function/line"}}
      ],
      "lineage_gaps": ["gap 1"],
      "recommended_fixes": ["fix 1"]
    }}

    Rules:
    - Prioritize semantic risk and lineage gaps.
    - Be concise and specific.
    """
).strip()


def render_prompt(template: str, **values: str) -> str:
    """Simple formatter for prompt templates."""
    return template.format(**values)
