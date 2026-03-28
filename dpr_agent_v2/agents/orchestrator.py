"""
agents/orchestrator.py
════════════════════════
Slim state machine. Only job: receive message → call correct handler → return response.
Zero business logic. Zero extraction. Zero benchmarks. Just routing.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from core.session_store import SessionStore
from agents.extractor import DEFAULT_MODEL
from agents.benchmark_engine import BenchmarkEngine
from agents.handlers.intake    import IntakeHandler
from agents.handlers.profile   import ProfileHandler
from agents.handlers.capital   import CapitalHandler
from agents.handlers.revenue   import RevenueHandler
from agents.handlers.costs     import CostsHandler
from agents.handlers.manpower  import ManpowerHandler
from agents.handlers.wc        import WCHandler
from agents.handlers.review    import ReviewHandler
from agents.handlers           import get_first_question
from core.industry_config      import get_profile, questions_to_skip


@dataclass
class OrchestratorResponse:
    message:          str
    ready_to_generate:bool = False
    section_completed:str  = ""


class Orchestrator:
    def __init__(self, api_key: str, model: str = DEFAULT_MODEL):
        self.api_key    = api_key
        self.model      = model
        self.store      = SessionStore()
        self.benchmarks : dict = {}
        self._reviewed  : bool = False

        # Initialise all handlers
        self.intake   = IntakeHandler(api_key, model)
        self.profile  = ProfileHandler(api_key, model)
        self.capital  = CapitalHandler(api_key, model)
        self.revenue  = RevenueHandler(api_key, model)
        self.costs    = CostsHandler(api_key, model)
        self.manpower = ManpowerHandler(api_key, model)
        self.wc       = WCHandler(api_key, model)
        self.review   = ReviewHandler(api_key, model)
        self.benchmark_engine = BenchmarkEngine(api_key, model)

    async def process(self, user_message: str) -> OrchestratorResponse:
        section = self.store.current_section

        # Skip sections not applicable to this industry
        profile       = get_profile(self.store.industry_code)
        skip_sections = questions_to_skip(profile)
        if section in skip_sections:
            self.store.current_section = self.store.next_incomplete_section()
            section = self.store.current_section

        # ── First message always goes through intake ─────────────────────────
        if section == "intake":
            msg = await self.intake.handle(user_message, self.store)
            return OrchestratorResponse(message=msg)

        # ── Section-by-section gap filling ───────────────────────────────────
        if section == "profile":
            msg = await self.profile.handle(user_message, self.store)
        elif section == "capital":
            msg = await self.capital.handle(user_message, self.store)
        elif section == "revenue":
            msg = await self.revenue.handle(user_message, self.store)
        elif section == "costs":
            msg = await self.costs.handle(user_message, self.store)
        elif section == "manpower":
            msg = await self.manpower.handle(user_message, self.store)
        elif section == "wc":
            msg = await self.wc.handle(user_message, self.store)

        # ── Review screen ─────────────────────────────────────────────────────
        elif section == "review":
            # First time we hit review: run benchmark engine, show review screen
            if not self._reviewed:
                self._reviewed = True
                self.benchmarks = await self.benchmark_engine.generate(self.store)
                self.benchmark_engine.apply_to_store(self.store, self.benchmarks)
                msg = self.review.build_review_screen(self.store, self.benchmarks)
                return OrchestratorResponse(message=msg)
            else:
                # User is responding to review screen
                msg, generate = await self.review.handle(
                    user_message, self.store, self.benchmarks)
                if generate:
                    return OrchestratorResponse(
                        message=msg, ready_to_generate=True)

        else:
            msg = "Something went wrong. Please refresh and try again."

        return OrchestratorResponse(message=msg)

    def get_store(self) -> SessionStore:
        return self.store
