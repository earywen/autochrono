# DECISIONS.md — Architecture Decision Records

> Log of significant technical and design decisions.

## Format

Each decision follows this structure:
- **Date**: When the decision was made
- **Context**: What prompted the decision
- **Decision**: What was decided
- **Consequences**: Impact of this decision

---

## ADR-001: Python + Tkinter for GUI

**Date**: 2026-02-04

**Context**: Need a simple GUI application that generates VBA code. Must be easy to distribute as .exe in enterprise environment.

**Decision**: Use Python 3.x with Tkinter for GUI, PyInstaller for packaging.

**Consequences**:
- ✅ Tkinter is included with Python, no external dependencies
- ✅ PyInstaller creates standalone .exe
- ✅ Native Windows look and feel
- ⚠️ Limited to basic UI components (acceptable for this use case)

---

## ADR-002: Manual VBA Import

**Date**: 2026-02-04

**Context**: Automatic VBA injection into Outlook requires admin rights which are not available.

**Decision**: Generate a .bas file with clear import instructions instead of automatic injection.

**Consequences**:
- ✅ No admin rights required
- ✅ No security alerts
- ⚠️ Users must follow manual import steps (mitigated by clear instructions)
