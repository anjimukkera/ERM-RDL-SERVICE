// DOCX_EDITABLE is the Windows Word page-locked renderer. DOCX_REFLOWABLE deliberately keeps the same
// resolved RDL content and native tables but allows Word to grow rows and preserve normal edit font sizes.
import { renderPagedEditableDocx } from './pagedDocx.js';

export function renderEditableDocx(model, request, config, tempDir, telemetry) {
  return renderPagedEditableDocx(model, request, config, tempDir, telemetry);
}

export function renderReflowableDocx(model, request, config, tempDir, telemetry) {
  return renderPagedEditableDocx(model, request, config, tempDir, telemetry, { reflowable: true });
}
