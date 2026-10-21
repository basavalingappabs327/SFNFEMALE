/**
 * Sterilization Certificate (Female applicant / SFN).
 *
 * Backend-calculated:
 *   {employeeageeng} → "46 years"
 *   {employeeage}    → "46 ವರ್ಷಗಳು"
 *   {dateofsfn}      → 1st of month after operation
 *
 * Requires oauth scope: https://www.googleapis.com/auth/script.external_request
 */
const TEMPLATE_DOC_ID = "1ONeKHrxr6j-NXcViiAZxsohM_TLQLYFYmkTVe5iV2ec";

function doGet() {
  return jsonOutput_({ success: true, message: "Sterilization Certificate generator is running." });
}

function doPost(e) {
  try {
    if (!TEMPLATE_DOC_ID || TEMPLATE_DOC_ID.indexOf("PASTE_") === 0) {
      throw new Error("Set TEMPLATE_DOC_ID in Code.gs to your converted Google Docs template ID.");
    }

    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const operationDate = parseDate_(data.dateofoperation);
    const dob = parseDate_(data.dobofemployee);
    if (!operationDate) throw new Error("A valid date of operation is required.");
    if (!dob) throw new Error("A valid employee date of birth is required.");
    if (dob.getTime() > operationDate.getTime()) {
      throw new Error("Employee date of birth cannot be after the date of operation.");
    }

    const ageYears = calculateAge_(dob, operationDate);
    data.dateofoperation = formatDate_(operationDate);
    data.dobofemployee = formatDate_(dob);
    data.dateofsfn = formatDate_(new Date(operationDate.getFullYear(), operationDate.getMonth() + 1, 1));
    data.employeeageeng = ageYears + " years";
    data.employeeage = ageYears + " ವರ್ಷಗಳು";
    data.today = data.today || formatDate_(new Date());

    const template = DriveApp.getFileById(TEMPLATE_DOC_ID);
    const copyName = "Sterilization_Certificate_" +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") +
      "_" + Utilities.getUuid().slice(0, 8);
    const copy = template.makeCopy(copyName);
    const copyId = copy.getId();

    try {
      const doc = DocumentApp.openById(copyId);
      const body = doc.getBody();
      const replacements = buildReplacements_(data);
      insertPageBreaks_(body);
      replaceInContainer_(body, replacements);
      const header = doc.getHeader();
      if (header) replaceInContainer_(header, replacements);
      const footer = doc.getFooter();
      if (footer) replaceInContainer_(footer, replacements);
      doc.saveAndClose();

      const pdfBlob = DriveApp.getFileById(copyId).getAs(MimeType.PDF).setName(copyName + ".pdf");
      const docxBlob = exportGoogleDocAsDocx_(copyId, copyName + ".docx");

      return jsonOutput_({
        success: true,
        pdfFileName: pdfBlob.getName(),
        pdfBase64: Utilities.base64Encode(pdfBlob.getBytes()),
        docxFileName: docxBlob.getName(),
        docxBase64: Utilities.base64Encode(docxBlob.getBytes()),
        employeeage: data.employeeage,
        employeeageeng: data.employeeageeng,
        dateofsfn: data.dateofsfn
      });
    } finally {
      try { DriveApp.getFileById(copyId).setTrashed(true); } catch (cleanupError) {}
    }
  } catch (err) {
    return jsonOutput_({ success: false, error: String(err && err.message ? err.message : err) });
  }
}

function exportGoogleDocAsDocx_(fileId, outputName) {
  const mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const url = "https://www.googleapis.com/drive/v3/files/" + encodeURIComponent(fileId) +
    "/export?mimeType=" + encodeURIComponent(mimeType);
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error("DOCX export failed (HTTP " + response.getResponseCode() + "): " + response.getContentText());
  }
  return response.getBlob().setContentType(mimeType).setName(outputName);
}

function insertPageBreaks_(body) {
  ["\\{pagebreak\\}", "\\{pagebreaK\\}"].forEach(function(pattern) {
    var found = body.findText(pattern);
    while (found) {
      try {
        var parent = found.getElement().getParent();
        if (parent && parent.getType && parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
          parent.asParagraph().appendPageBreak();
        } else if (parent && parent.appendPageBreak) {
          parent.appendPageBreak();
        }
      } catch (e) {}
      found = body.findText(pattern, found);
    }
  });
}

function buildReplacements_(d) {
  const result = {};
  const keys = [
    "doctor", "operation", "employee", "spouse", "designation",
    "instiution", "institution", "institutionplace", "instituitionplace",
    "taluk", "district", "hospital",
    "dateofoperation", "dateofsfn", "dobofemployee",
    "employeeage", "employeeageeng", "spouseage", "spousekannada", "spouseagekannada",
    "noofchildren", "1age", "2age",
    "employeekannada", "designationkannada",
    "institutionkannada", "institutionplacekannada", "instiutionplacekannada",
    "hospitalkannada", "noofchildrenkannada",
    "talukkannada", "districtkannada",
    "basic", "scale", "today",
    "pagebreak", "pagebreaK"
  ];
  keys.forEach(function(key) {
    result[key] = d[key] == null ? "" : String(d[key]);
  });
  result.institution = d.instiution || d.institution || "";
  result.instiution = d.instiution || d.institution || "";
  result.instituitionplace = d.institutionplace || d.instituitionplace || "";
  result.institutionplace = d.institutionplace || "";
  result.instiutionplacekannada = d.institutionplacekannada || d.instiutionplacekannada || "";
  result.institutionplacekannada = d.institutionplacekannada || d.instiutionplacekannada || "";
  result.pagebreak = "";
  result.pagebreaK = "";
  return result;
}

function replaceInContainer_(container, replacements) {
  Object.keys(replacements).forEach(function(key) {
    const pattern = "\\{" + key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\}";
    container.replaceText(pattern, String(replacements[key] == null ? "" : replacements[key]));
  });
}

function parseDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  const text = String(value).trim();
  let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) {
    const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
    if (date.getFullYear() === Number(match[3]) && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[1])) return date;
    return null;
  }
  match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    if (date.getFullYear() === Number(match[1]) && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[3])) return date;
  }
  return null;
}

function calculateAge_(dob, operationDate) {
  let age = operationDate.getFullYear() - dob.getFullYear();
  const birthdayThisYear = new Date(operationDate.getFullYear(), dob.getMonth(), dob.getDate());
  if (operationDate < birthdayThisYear) age--;
  return age;
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function jsonOutput_(object) {
  return ContentService.createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}