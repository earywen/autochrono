# DECISIONS.md — Architecture Decision Records

> Log of significant technical and design decisions.

## Format

Each decision follows this structure:
- **Date**: When the decision was made
- **Context**: What prompted the decision
- **Decision**: What was decided
- **Consequences**: Impact of this decision

---

## ADR-001: Python + CustomTkinter for GUI

**Date**: 2026-02-04

**Context**: Need a GUI application that generates VBA code. Must be easy to distribute as .exe in enterprise environment. User wants a modern, professional-looking interface.

**Decision**: Use Python 3.x with CustomTkinter for GUI, PyInstaller for packaging.

**Consequences**:
- ✅ Modern UI with dark mode and rounded corners
- ✅ Professional look for enterprise environment
- ✅ PyInstaller creates standalone .exe
- ⚠️ Requires customtkinter dependency (mitigated by PyInstaller bundling)

---

## ADR-002: Manual VBA Import

**Date**: 2026-02-04

**Context**: Automatic VBA injection into Outlook requires admin rights which are not available.

**Decision**: Generate a .bas file with clear import instructions instead of automatic injection.

**Consequences**:
- ✅ No admin rights required
- ✅ No security alerts
- ⚠️ Users must follow manual import steps (mitigated by clear instructions)
