const SHEET_ID = "XXX";
const WEB_APP_URL = "XXX";
const FOLDER_ID = "XXX";
const ADMIN_PASSWORD = "CHANGE_ME_BEFORE_SETUP";

var APP_VERSION = "1.0.0-gas";
var SESSION_HOURS = 8;
var MAX_UPLOAD_BYTES = 10 * 1024 * 1024;
var PASSWORD_ROUNDS = 1200;
var APPROVED = "APPROVED";
var ACTIVE = "ACTIVE";
var ALLOWED_ROLES = ["Super Admin", "Admin", "Credit Editor", "Staff"];
var ALLOWED_UPLOAD_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "image/jpeg",
  "image/png",
  "image/webp"
];

var SHEET_SCHEMAS = {
  Settings: ["key","value","updatedAt"],
  Users: ["id","username","passwordHash","passwordSalt","name","branch","role","status","mustChangePassword","createdAt","updatedAt"],
  Sessions: ["tokenHash","userId","expiresAt","revoked","createdAt","lastAccess"],
  Occupations: ["occupationId","occupationCode","occupationName","occupationGroup","description","characteristics","requiredDocuments","supportingDocuments","interviewQuestions","incomeInputs","incomeMethod","incomeFormula","expenseMethod","expenseRate","expenseConditions","warning","riskPoints","calculationNotes","sourceDocument","sourcePage","version","effectiveDate","endDate","approvalStatus","status","updatedBy","updatedAt"],
  Purposes: ["purposeId","purposeCode","purposeName","description","requiredCollateralDocuments","optionalDocuments","specialConditions","warning","source","sourcePage","version","effectiveDate","endDate","approvalStatus","status","updatedBy","updatedAt"],
  IncomeRules: ["id","occupationCode","ruleType","inputFields","formulaType","formulaConfig","condition","source","sourcePage","version","effectiveDate","endDate","approvalStatus","status","updatedBy","updatedAt"],
  ExpenseRules: ["id","occupationCode","expenseType","rate","formula","condition","source","sourcePage","version","effectiveDate","endDate","approvalStatus","status","updatedBy","updatedAt"],
  DSRRules: ["id","ruleName","customerType","productGroup","incomeMin","incomeMax","maxDSR","condition","source","sourcePage","effectiveDate","endDate","version","approvalStatus","status","updatedBy","updatedAt"],
  Products: ["productId","productCode","productName","customerType","occupationCondition","purposeCondition","incomeCondition","ageCondition","loanCondition","specialCondition","effectiveDate","expiryDate","source","sourcePage","version","approvalStatus","status","updatedBy","updatedAt"],
  ProductRules: ["id","productCode","field","operator","value","condition","message","source","sourcePage","version","effectiveDate","endDate","approvalStatus","status","updatedBy","updatedAt"],
  ChecklistRules: ["id","occupationCode","purposeCode","documentGroup","documentName","required","condition","customerText","staffNote","source","sourcePage","version","effectiveDate","endDate","approvalStatus","order","status","updatedBy","updatedAt"],
  RiskAlerts: ["id","riskCode","occupationCode","conditionType","field","operator","value","severity","message","recommendation","source","sourcePage","version","effectiveDate","endDate","approvalStatus","status","updatedBy","updatedAt"],
  KnowledgeBase: ["id","category","occupationCode","title","content","keywords","sourceDocument","sourcePage","version","effectiveDate","endDate","approvalStatus","status","createdAt","updatedAt","updatedBy"],
  QALibrary: ["id","question","keywords","answer","sourceDocument","sourcePage","version","effectiveDate","endDate","approvalStatus","status","updatedAt","updatedBy"],
  CaseLogs: ["id","caseRef","userId","branch","occupationCode","purposeCode","module","resultSummary","createdAt"],
  UsageLogs: ["id","sessionId","userId","branch","module","occupationCode","purposeCode","startTime","endTime","durationSeconds","success","createdAt"],
  Feedback: ["id","userId","branch","easeOfUse","informationClarity","accuracyPerception","timeSaving","checklistUsefulness","overallSatisfaction","comment","createdAt"],
  Files: ["id","fileId","fileUrl","fileName","fileType","fileSize","folder","resourceKey","createdAt","uploadedBy"],
  Logs: ["id","username","action","module","recordId","detail","oldValue","newValue","timestamp"]
};

var OCCUPATION_CATALOG = [
  ["OCC001","รับเหมาก่อสร้าง"],["OCC002","ขายเสื้อผ้าสำเร็จรูป"],["OCC003","เสื้อผ้าตัดเย็บ"],["OCC004","ทำฟาร์ม"],
  ["OCC005","นายหน้า"],["OCC006","ทำพืชไร่"],["OCC007","สวนยาง/สวนเกษตร"],["OCC008","ค้าขาย"],
  ["OCC009","ขายประกัน"],["OCC010","ขายอาหาร"],["OCC011","ที่ปรึกษา"],["OCC012","จัดสรรที่ดิน"],
  ["OCC013","ทนายความ"],["OCC014","รับจ้างกลึง/เชื่อม"],["OCC015","ศิลปินงานแสดง"],["OCC016","ขายของชำ"],
  ["OCC017","ขายสินค้าเงินผ่อน"],["OCC018","จำหน่ายวัสดุ/อุปกรณ์ก่อสร้าง"],["OCC019","ให้เช่ารถ"],["OCC020","ตัดเย็บเสื้อผ้าโหล"],
  ["OCC021","ขับรถรับจ้างทั่วไป"],["OCC022","รับจ้างผลิตสื่อโฆษณา"],["OCC023","เทรนเนอร์"],["OCC024","นักกีฬาอาชีพ"],
  ["OCC025","สื่อและบริการออนไลน์"],["OCC026","Delivery"],["OCC027","พระเครื่อง/วัตถุมงคล"],["OCC028","ขายสินค้าออนไลน์"]
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Homie Credit Assistant")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1");
}

function setupSystem() {
  try {
    validateConfig_();
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var root = DriveApp.getFolderById(FOLDER_ID);
    var createdSheets = [];
    var name;
    for (name in SHEET_SCHEMAS) {
      if (SHEET_SCHEMAS.hasOwnProperty(name)) {
        if (ensureSheet_(ss, name, SHEET_SCHEMAS[name])) createdSheets.push(name);
      }
    }
    var folders = ensureDriveFolders_(root);
    seedSettings_();
    seedOccupationCatalog_();
    seedAdmin_();
    clearRuntimeCaches_();
    var schemaResult = validateSchema_();
    writeLog_("system", "SETUP", "System", "", "setupSystem completed", "", JSON.stringify(schemaResult));
    return ok_({version: APP_VERSION,spreadsheetName: ss.getName(),createdSheets: createdSheets,folders: folders,schema: schemaResult,note: "นำเข้า Homie V7 Master ที่มี Source/Version/Approval ก่อนเปิด Runtime Rule"});
  } catch (error) {
    serverError_("setupSystem", error);
    return fail_(friendlyError_(error));
  }
}

function login(payload) {
  try {
    payload = payload || {};
    var username = cleanText_(payload.username, 80).toLowerCase();
    var password = String(payload.password || "");
    if (!username || !password) return fail_("กรุณากรอก Username และ Password");
    var users = readSheetObjects_("Users");
    var user = findOne_(users, "username", username);
    if (!user || String(user.status).toUpperCase() !== ACTIVE) {
      writeLog_(username || "unknown", "FAILED_LOGIN", "Auth", "", "Invalid username or inactive user", "", "");
      return fail_("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง");
    }
    var computed = hashPassword_(password, user.passwordSalt);
    if (!safeEqual_(computed, user.passwordHash)) {
      writeLog_(username, "FAILED_LOGIN", "Auth", user.id, "Invalid password", "", "");
      return fail_("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง");
    }
    var session = createSession_(user);
    writeLog_(username, "LOGIN", "Auth", user.id, "Login success", "", "");
    return ok_({token: session.token,expiresAt: session.expiresAt,user: publicUser_(user),mustChangePassword: String(user.mustChangePassword).toUpperCase() === "TRUE",appVersion: APP_VERSION});
  } catch (error) {
    serverError_("login", error);
    return fail_("ไม่สามารถเข้าสู่ระบบได้");
  }
}

function logout(payload) {
  try {
    payload = payload || {};
    var token = String(payload.token || "");
    if (!token) return ok_(true);
    var tokenHash = sha256_(token);
    var sessions = readSheetObjects_("Sessions");
    var i;
    for (i = 0; i < sessions.length; i++) {
      if (sessions[i].tokenHash === tokenHash && String(sessions[i].revoked).toUpperCase() !== "TRUE") {
        sessions[i].revoked = "TRUE";
        updateRowByKey_("Sessions", "tokenHash", tokenHash, sessions[i]);
        var u = getUserById_(sessions[i].userId);
        writeLog_(u ? u.username : "unknown", "LOGOUT", "Auth", sessions[i].userId, "Logout", "", "");
      }
    }
    return ok_(true);
  } catch (error) {
    serverError_("logout", error);
    return fail_("ออกจากระบบไม่สำเร็จ");
  }
}

function checkSession(payload) {
  try {
    var auth = requireSession_(payload && payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    return ok_({ user: publicUser_(auth.user), expiresAt: auth.session.expiresAt, appVersion: APP_VERSION });
  } catch (error) {
    serverError_("checkSession", error);
    return fail_("Session ไม่ถูกต้อง");
  }
}

function changePassword(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var oldPassword = String(payload.oldPassword || "");
    var newPassword = String(payload.newPassword || "");
    if (!safeEqual_(hashPassword_(oldPassword, auth.user.passwordSalt), auth.user.passwordHash)) return fail_("รหัสผ่านเดิมไม่ถูกต้อง");
    var valid = validatePassword_(newPassword);
    if (!valid.ok) return fail_(valid.message);
    var salt = randomToken_().substring(0, 40);
    var updated = clone_(auth.user);
    updated.passwordSalt = salt;
    updated.passwordHash = hashPassword_(newPassword, salt);
    updated.mustChangePassword = "FALSE";
    updated.updatedAt = nowIso_();
    updateRowByKey_("Users", "id", auth.user.id, updated);
    revokeUserSessions_(auth.user.id, payload.token);
    writeLog_(auth.user.username, "CHANGE_PASSWORD", "Users", auth.user.id, "Password changed", "", "");
    return ok_(true);
  } catch (error) {
    serverError_("changePassword", error);
    return fail_("เปลี่ยนรหัสผ่านไม่สำเร็จ");
  }
}

function getHomeDashboard(payload) {
  try {
    var auth = requireSession_(payload && payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var usage = readSheetObjects_("UsageLogs");
    var feedback = readSheetObjects_("Feedback");
    var cases = readSheetObjects_("CaseLogs");
    var userId = auth.user.id;
    var userUsage = filterBy_(usage, function(r){ return r.userId === userId; });
    var userFeedback = filterBy_(feedback, function(r){ return r.userId === userId; });
    var userCases = filterBy_(cases, function(r){ return r.userId === userId; });
    var occupationCounts = countBy_(userUsage, "occupationCode");
    return ok_({sessions: countUserSessions_(userId),cases: userCases.length,topOccupation: topKey_(occupationCounts) || "-",checklistCount: countWhere_(userUsage, function(r){ return r.module === "Smart Checklist" && String(r.success).toUpperCase() === "TRUE"; }),estimatedMinutes: Math.round(sumField_(userUsage, "durationSeconds") / 60),feedbackScore: averageField_(userFeedback, "overallSatisfaction"),disclaimer: getSettingValue_("DISCLAIMER")});
  } catch (error) { serverError_("getHomeDashboard", error); return fail_("โหลด Dashboard ไม่สำเร็จ"); }
}

function getOccupations(payload) {
  try {
    var auth = requireSession_(payload && payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var rows = readSheetObjects_("Occupations");
    if (!isAdminRole_(auth.user.role)) rows = filterBy_(rows, function(r){ return isRuntimeMaster_(r); });
    return ok_(rows);
  } catch (error) { serverError_("getOccupations", error); return fail_("โหลดอาชีพไม่สำเร็จ"); }
}

function getOccupationById(payload) {
  try {
    var auth = requireSession_(payload && payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var key = cleanText_(payload.id || payload.occupationCode, 80);
    var rows = readSheetObjects_("Occupations");
    var row = findOne_(rows, "occupationId", key) || findOne_(rows, "occupationCode", key);
    if (!row) return fail_("ไม่พบอาชีพ");
    if (!isAdminRole_(auth.user.role) && !isRuntimeMaster_(row)) return fail_(approvedMissingMessage_());
    return ok_(row);
  } catch (error) { serverError_("getOccupationById", error); return fail_("โหลดข้อมูลอาชีพไม่สำเร็จ"); }
}

function saveOccupation(payload) { return saveMaster_(payload, "Occupations", "occupationId", "OCCREC_", ["Super Admin","Admin","Credit Editor"], false); }
function deleteOccupation(payload) { return deleteMaster_(payload, "Occupations", "occupationId", ["Super Admin","Admin"]); }
function getPurposes(payload) { return getMaster_(payload, "Purposes", ["Super Admin","Admin","Credit Editor","Staff"], true); }
function savePurpose(payload) { return saveMaster_(payload, "Purposes", "purposeId", "PURREC_", ["Super Admin","Admin","Credit Editor"], false); }
function deletePurpose(payload) { return deleteMaster_(payload, "Purposes", "purposeId", ["Super Admin","Admin"]); }
function getIncomeRules(payload) { return getMaster_(payload, "IncomeRules", ["Super Admin","Admin","Credit Editor"], false); }
function saveIncomeRule(payload) { return saveMaster_(payload, "IncomeRules", "id", "INCRULE_", ["Super Admin","Admin","Credit Editor"], true); }
function deleteIncomeRule(payload) { return deleteMaster_(payload, "IncomeRules", "id", ["Super Admin","Admin"]); }
function getExpenseRules(payload) { return getMaster_(payload, "ExpenseRules", ["Super Admin","Admin","Credit Editor"], false); }
function saveExpenseRule(payload) { return saveMaster_(payload, "ExpenseRules", "id", "EXPRULE_", ["Super Admin","Admin","Credit Editor"], true); }
function deleteExpenseRule(payload) { return deleteMaster_(payload, "ExpenseRules", "id", ["Super Admin","Admin"]); }
function getDSRRules(payload) { return getMaster_(payload, "DSRRules", ["Super Admin","Admin"], false); }
function saveDSRRule(payload) { return saveMaster_(payload, "DSRRules", "id", "DSRRULE_", ["Super Admin","Admin"], true); }
function deleteDSRRule(payload) { return deleteMaster_(payload, "DSRRules", "id", ["Super Admin","Admin"]); }
function getProducts(payload) { return getMaster_(payload, "Products", ["Super Admin","Admin","Credit Editor","Staff"], true); }
function saveProduct(payload) { return saveMaster_(payload, "Products", "productId", "PROD_", ["Super Admin","Admin"], true); }
function deleteProduct(payload) { return deleteMaster_(payload, "Products", "productId", ["Super Admin","Admin"]); }
function getProductRules(payload) { return getMaster_(payload, "ProductRules", ["Super Admin","Admin"], false); }
function saveProductRule(payload) { return saveMaster_(payload, "ProductRules", "id", "PRULE_", ["Super Admin","Admin"], true); }
function deleteProductRule(payload) { return deleteMaster_(payload, "ProductRules", "id", ["Super Admin","Admin"]); }
function getChecklistRules(payload) { return getMaster_(payload, "ChecklistRules", ["Super Admin","Admin","Credit Editor"], false); }
function saveChecklistRule(payload) { return saveMaster_(payload, "ChecklistRules", "id", "CKRULE_", ["Super Admin","Admin","Credit Editor"], true); }
function deleteChecklistRule(payload) { return deleteMaster_(payload, "ChecklistRules", "id", ["Super Admin","Admin"]); }
function getRiskAlerts(payload) { return getMaster_(payload, "RiskAlerts", ["Super Admin","Admin","Credit Editor"], false); }
function saveRiskAlert(payload) { return saveMaster_(payload, "RiskAlerts", "id", "RISK_", ["Super Admin","Admin","Credit Editor"], true); }
function deleteRiskAlert(payload) { return deleteMaster_(payload, "RiskAlerts", "id", ["Super Admin","Admin"]); }
function getKnowledge(payload) { return getMaster_(payload, "KnowledgeBase", ["Super Admin","Admin","Credit Editor","Staff"], true); }
function getKnowledgeById(payload) { return getKnowledgeItem_(payload); }
function saveKnowledge(payload) { return saveMaster_(payload, "KnowledgeBase", "id", "KB_", ["Super Admin","Admin","Credit Editor"], true); }
function deleteKnowledge(payload) { return deleteMaster_(payload, "KnowledgeBase", "id", ["Super Admin","Admin"]); }
function getQALibrary(payload) { return getMaster_(payload, "QALibrary", ["Super Admin","Admin","Credit Editor"], false); }
function saveQA(payload) { return saveMaster_(payload, "QALibrary", "id", "QA_", ["Super Admin","Admin","Credit Editor"], true); }
function deleteQA(payload) { return deleteMaster_(payload, "QALibrary", "id", ["Super Admin","Admin"]); }

function getIncomeContext(payload) {
  try {
    var auth = requireSession_(payload && payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var occupationCode = cleanText_(payload.occupationCode, 50);
    var occ = findRuntimeOccupation_(occupationCode);
    if (!occ) return fail_(approvedMissingMessage_());
    var incomeRules = runtimeRules_("IncomeRules", function(r){ return r.occupationCode === occupationCode; });
    var expenseRules = runtimeRules_("ExpenseRules", function(r){ return r.occupationCode === occupationCode; });
    if (!incomeRules.length) return fail_(approvedMissingMessage_());
    if (incomeRules.length > 1) return fail_("พบ Rule รายได้ Active ที่ขัดแย้งกัน กรุณาแจ้งผู้ดูแลระบบ");
    if (expenseRules.length > 1) return fail_("พบ Rule ค่าใช้จ่าย Active ที่ขัดแย้งกัน กรุณาแจ้งผู้ดูแลระบบ");
    return ok_({occupation: occ,inputFields: safeJson_(incomeRules[0].inputFields, []),incomeRule: sourceView_(incomeRules[0]),expenseRule: expenseRules.length ? sourceView_(expenseRules[0]) : null,requiredDocuments: safeJson_(occ.requiredDocuments, []),supportingDocuments: safeJson_(occ.supportingDocuments, []),interviewQuestions: safeJson_(occ.interviewQuestions, []),warnings: safeJson_(occ.riskPoints, []).concat(occ.warning ? [occ.warning] : [])});
  } catch (error) { serverError_("getIncomeContext", error); return fail_("โหลด Rule สำหรับอาชีพไม่สำเร็จ"); }
}

function calculateIncome(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var occupationCode = cleanText_(payload.occupationCode, 50);
    var occ = findRuntimeOccupation_(occupationCode);
    if (!occ) return fail_(approvedMissingMessage_());
    var incomeRules = runtimeRules_("IncomeRules", function(r){ return r.occupationCode === occupationCode; });
    var expenseRules = runtimeRules_("ExpenseRules", function(r){ return r.occupationCode === occupationCode; });
    if (incomeRules.length !== 1) return fail_(incomeRules.length === 0 ? "ไม่สามารถประเมินได้ เนื่องจากข้อมูล Rule รายได้ไม่สมบูรณ์" : "พบ Rule รายได้ Active มากกว่า 1 รายการ ระบบหยุดการคำนวณเพื่อป้องกัน Conflict");
    if (expenseRules.length > 1) return fail_("พบ Rule ค่าใช้จ่าย Active มากกว่า 1 รายการ ระบบหยุดการคำนวณเพื่อป้องกัน Conflict");
    var inputs = payload.inputs && typeof payload.inputs === "object" ? payload.inputs : {};
    var validation = validateDynamicInputs_(safeJson_(incomeRules[0].inputFields, []), inputs);
    if (!validation.ok) return fail_(validation.message);
    var gross = computeFormula_(incomeRules[0], inputs);
    if (!gross.ok) return fail_(gross.message);
    var expenseResult = computeExpense_(expenseRules.length ? expenseRules[0] : null, inputs, gross.value);
    if (!expenseResult.ok) return fail_(expenseResult.message);
    var net = gross.value - expenseResult.value;
    if (net < 0) return fail_("ผลรายได้สุทธิติดลบ กรุณาตรวจสอบข้อมูลและ Rule");
    var result = {occupationCode: occupationCode,grossIncome: round2_(gross.value),expense: round2_(expenseResult.value),netIncome: round2_(net),method: incomeRules[0].formulaType,calculationDetail: gross.detail.concat(expenseResult.detail),warnings: collectWarnings_(occ, incomeRules[0], expenseRules.length ? expenseRules[0] : null),source: incomeRules[0].source,sourcePage: incomeRules[0].sourcePage,sourceDetail: [sourceView_(incomeRules[0]), expenseRules.length ? sourceView_(expenseRules[0]) : null]};
    recordUsageInternal_(auth, "Income Assessment", occupationCode, "", true, payload.startTime);
    return ok_(result);
  } catch (error) { serverError_("calculateIncome", error); return fail_("ประเมินรายได้ไม่สำเร็จ"); }
}

function calculateDSR(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var netIncome = numberOrNull_(payload.netIncome);
    var existingDebt = numberOrNull_(payload.existingDebt);
    var proposedInstallment = numberOrNull_(payload.proposedInstallment);
    if (netIncome === null || netIncome <= 0) return fail_("รายได้สุทธิต้องมากกว่า 0");
    if (existingDebt === null || existingDebt < 0 || proposedInstallment === null || proposedInstallment < 0) return fail_("ภาระหนี้ต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป");
    var customerType = cleanText_(payload.customerType, 100);
    var productGroup = cleanText_(payload.productGroup, 100);
    var rules = runtimeRules_("DSRRules", function(r){var min = numberOrNull_(r.incomeMin);var max = numberOrNull_(r.incomeMax);var inRange = (min === null || netIncome >= min) && (max === null || max === 0 || netIncome <= max);var customerOk = !r.customerType || r.customerType === "ALL" || r.customerType === customerType;var productOk = !r.productGroup || r.productGroup === "ALL" || r.productGroup === productGroup;return inRange && customerOk && productOk;});
    if (rules.length === 0) return fail_("ไม่พบ DSR Rule ที่ได้รับอนุมัติและตรงเงื่อนไข");
    if (rules.length > 1) return fail_("พบ DSR Rule ที่ขัดแย้งกัน ระบบหยุดการคำนวณ กรุณาแจ้งผู้ดูแลระบบ");
    var maxDSR = numberOrNull_(rules[0].maxDSR);
    if (maxDSR === null || maxDSR <= 0) return fail_("DSR Rule ไม่มีค่า maxDSR ที่ใช้งานได้");
    var totalDebt = existingDebt + proposedInstallment;
    var dsr = totalDebt / netIncome * 100;
    var text = dsr <= maxDSR ? "อยู่ในช่วงเกณฑ์ที่กำหนด" : "เกินเกณฑ์เบื้องต้น";
    if (Math.abs(dsr - maxDSR) <= 2) text = "ควรตรวจสอบเพิ่มเติม";
    return ok_({netIncome: round2_(netIncome),existingDebt: round2_(existingDebt),proposedInstallment: round2_(proposedInstallment),totalDebt: round2_(totalDebt),dsr: round2_(dsr),applicableRule: rules[0].ruleName,maxDSR: maxDSR,statusText: text,warning: "ผล DSR เป็นข้อมูลเบื้องต้นสำหรับช่วยพนักงานเตรียมเคสเท่านั้น",source: sourceView_(rules[0])});
  } catch (error) { serverError_("calculateDSR", error); return fail_("คำนวณ DSR ไม่สำเร็จ"); }
}

function estimateLoan(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var installment = numberOrNull_(payload.installment);
    if (installment === null || installment <= 0) return fail_("กรุณาระบุความสามารถผ่อนที่มากกว่า 0");
    var config = getApprovedLoanEstimateConfig_();
    if (!config.ok) return fail_(config.message);
    var annualRate = numberOrNull_(config.data.annualInterestRate);
    var termMonths = numberOrNull_(payload.termMonths);
    if (termMonths === null || termMonths <= 0) termMonths = numberOrNull_(config.data.defaultTermMonths);
    var maxTerm = numberOrNull_(config.data.maxTermMonths);
    if (maxTerm !== null && termMonths > maxTerm) termMonths = maxTerm;
    if (annualRate === null || annualRate < 0 || !termMonths) return fail_("Loan Estimate Rule ไม่สมบูรณ์");
    var monthlyRate = annualRate / 100 / 12;
    var principal = monthlyRate === 0 ? installment * termMonths : installment * (1 - Math.pow(1 + monthlyRate, -termMonths)) / monthlyRate;
    return ok_({installment: round2_(installment),termMonths: termMonths,annualInterestRate: annualRate,estimatedLoan: round2_(principal),source: config.data.source,sourcePage: config.data.sourcePage,version: config.data.version,disclaimer: "เป็นการประมาณเพื่อประกอบการเตรียมเคส ไม่ใช่ผลอนุมัติสินเชื่อ"});
  } catch (error) { serverError_("estimateLoan", error); return fail_("ประมาณวงเงินไม่สำเร็จ"); }
}

function matchProducts(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var context = payload.context && typeof payload.context === "object" ? payload.context : {};
    var products = runtimeRules_("Products", function(){ return true; });
    var rules = runtimeRules_("ProductRules", function(){ return true; });
    var matches = [];
    var i;
    for (i = 0; i < products.length; i++) {
      var product = products[i];
      var productRules = filterBy_(rules, function(r){ return r.productCode === product.productCode; });
      if (!productRules.length) continue;
      var passed = true;
      var reasons = [];
      var checks = [];
      var j;
      for (j = 0; j < productRules.length; j++) {
        var ev = evaluateCondition_(context[productRules[j].field], productRules[j].operator, productRules[j].value);
        checks.push({ field: productRules[j].field, passed: ev, message: productRules[j].message, source: sourceView_(productRules[j]) });
        if (!ev) passed = false;
        if (ev && productRules[j].message) reasons.push(productRules[j].message);
      }
      if (passed) matches.push({productCode: product.productCode,productName: product.productName,matchText: "ผลิตภัณฑ์ที่อาจเข้าเงื่อนไขเบื้องต้น",reasons: reasons,conditionsToVerify: checks,source: sourceView_(product),demo: isDemoRecord_(product)});
    }
    return ok_({ matches: matches, disclaimer: "ผลการจับคู่ไม่ใช่การยืนยันสิทธิ์หรือผลอนุมัติ" });
  } catch (error) { serverError_("matchProducts", error); return fail_("จับคู่ผลิตภัณฑ์ไม่สำเร็จ"); }
}

function generateChecklist(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var occ = cleanText_(payload.occupationCode, 50);
    var pur = cleanText_(payload.purposeCode, 50);
    if (!occ || !pur) return fail_("กรุณาเลือกอาชีพและวัตถุประสงค์กู้");
    var occMaster = findRuntimeOccupation_(occ);
    var purpose = findRuntimePurpose_(pur);
    if (!occMaster || !purpose) return fail_(approvedMissingMessage_());
    var rules = runtimeRules_("ChecklistRules", function(r){var occOk = !r.occupationCode || r.occupationCode === "ALL" || r.occupationCode === occ;var purOk = !r.purposeCode || r.purposeCode === "ALL" || r.purposeCode === pur;return occOk && purOk;});
    if (!rules.length) return fail_("ไม่พบ Checklist Rule ที่ได้รับอนุมัติสำหรับอาชีพและวัตถุประสงค์นี้");
    rules.sort(function(a,b){ return Number(a.order || 0) - Number(b.order || 0); });
    var seen = {};
    var groups = {};
    var i;
    for (i = 0; i < rules.length; i++) {
      var key = String(rules[i].documentGroup || "") + "|" + String(rules[i].documentName || "").toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      var group = rules[i].documentGroup || "Additional Documents";
      if (!groups[group]) groups[group] = [];
      groups[group].push({id: rules[i].id,documentName: rules[i].documentName,required: String(rules[i].required).toUpperCase() === "TRUE",condition: rules[i].condition,customerText: rules[i].customerText || rules[i].documentName,staffNote: rules[i].staffNote,checked: false,source: sourceView_(rules[i]),demo: isDemoRecord_(rules[i])});
    }
    recordUsageInternal_(auth, "Smart Checklist", occ, pur, true, payload.startTime);
    return ok_({occupation: { code: occMaster.occupationCode, name: occMaster.occupationName },purpose: { code: purpose.purposeCode, name: purpose.purposeName },groups: groups,generatedAt: nowIso_(),disclaimer: getSettingValue_("DISCLAIMER")});
  } catch (error) { serverError_("generateChecklist", error); return fail_("สร้าง Checklist ไม่สำเร็จ"); }
}

function exportChecklistPdf(payload) {
  var tempDoc = null;
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var checklist = payload.checklist || {};
    var root = DriveApp.getFolderById(FOLDER_ID);
    var folder = getOrCreateChildFolder_(root, "Homie");
    folder = getOrCreateChildFolder_(folder, "Exported_Checklists");
    tempDoc = DocumentApp.create("HOMIE_TEMP_" + Utilities.getUuid());
    var body = tempDoc.getBody();
    body.appendParagraph("Homie Credit Assistant").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("Smart Checklist อาชีพอิสระ");
    body.appendParagraph("อาชีพ: " + cleanText_(checklist.occupationName || checklist.occupationCode, 300));
    body.appendParagraph("วัตถุประสงค์: " + cleanText_(checklist.purposeName || checklist.purposeCode, 300));
    var groups = checklist.groups && typeof checklist.groups === "object" ? checklist.groups : {};
    var groupName;
    for (groupName in groups) {
      if (!groups.hasOwnProperty(groupName)) continue;
      body.appendParagraph(groupName).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      var items = safeArray_(groups[groupName]);
      var i;
      for (i = 0; i < items.length; i++) body.appendListItem((items[i].checked ? "☑ " : "☐ ") + cleanText_(items[i].customerText || items[i].documentName, 500));
    }
    body.appendParagraph(getSettingValue_("DISCLAIMER"));
    tempDoc.saveAndClose();
    var tempFile = DriveApp.getFileById(tempDoc.getId());
    var pdfBlob = tempFile.getAs(MimeType.PDF).setName("Homie_Checklist_" + timestampForFile_() + ".pdf");
    var pdfFile = folder.createFile(pdfBlob);
    tempFile.setTrashed(true);
    tempDoc = null;
    writeFileRecord_(pdfFile, "Exported_Checklists", auth.user.username);
    writeLog_(auth.user.username, "EXPORT_PDF", "Checklist", pdfFile.getId(), "Checklist PDF exported", "", "");
    return ok_({ fileId: pdfFile.getId(), fileUrl: pdfFile.getUrl(), fileName: pdfFile.getName() });
  } catch (error) {
    try { if (tempDoc) DriveApp.getFileById(tempDoc.getId()).setTrashed(true); } catch (ignore) {}
    serverError_("exportChecklistPdf", error);
    return fail_("Export PDF ไม่สำเร็จ");
  }
}

function evaluateRiskAlerts(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var context = payload.context && typeof payload.context === "object" ? payload.context : {};
    var occ = cleanText_(payload.occupationCode, 50);
    var rules = runtimeRules_("RiskAlerts", function(r){ return !r.occupationCode || r.occupationCode === "ALL" || r.occupationCode === occ; });
    var alerts = [];
    var i;
    for (i = 0; i < rules.length; i++) if (evaluateCondition_(context[rules[i].field], rules[i].operator, rules[i].value)) alerts.push({riskCode: rules[i].riskCode,severity: rules[i].severity,message: rules[i].message,recommendation: rules[i].recommendation,source: sourceView_(rules[i]),demo: isDemoRecord_(rules[i])});
    return ok_({ alerts: alerts, disclaimer: "Risk Alert เป็นจุดเตือนเพื่อให้พนักงานตรวจสอบเพิ่มเติม ไม่ใช่ Credit Decision" });
  } catch (error) { serverError_("evaluateRiskAlerts", error); return fail_("ประเมิน Risk Alert ไม่สำเร็จ"); }
}

function searchKnowledge(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var keyword = cleanText_(payload.keyword, 200).toLowerCase();
    var category = cleanText_(payload.category, 100);
    var occ = cleanText_(payload.occupationCode, 50);
    var rows = runtimeRules_("KnowledgeBase", function(r){var categoryOk = !category || r.category === category;var occOk = !occ || !r.occupationCode || r.occupationCode === "ALL" || r.occupationCode === occ;if (!categoryOk || !occOk) return false;if (!keyword) return true;var hay = (String(r.title || "") + " " + String(r.content || "") + " " + String(r.keywords || "")).toLowerCase();return hay.indexOf(keyword) >= 0;});
    return ok_(rows);
  } catch (error) { serverError_("searchKnowledge", error); return fail_("ค้น Knowledge Base ไม่สำเร็จ"); }
}

function askHomie(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var question = cleanText_(payload.question, 1000);
    if (!question) return fail_("กรุณาระบุคำถาม");
    if (isApprovalQuestion_(question)) return ok_({answer: "Homie ไม่สามารถตัดสินผลอนุมัติสินเชื่อแทนผู้มีอำนาจได้ สามารถช่วยตรวจสอบหลักเกณฑ์ เอกสาร และข้อมูลประกอบการพิจารณาได้",sourceDocument: "Homie Governance Guardrail",sourcePage: "-",version: APP_VERSION,lastUpdated: nowIso_(),guardrail: true});
    var qa = runtimeRules_("QALibrary", function(){ return true; });
    var best = bestTextMatch_(qa, question, ["question","keywords","answer"]);
    if (best && best.score > 0) {
      recordUsageInternal_(auth, "Q&A", "", "", true, payload.startTime);
      return ok_({ answer: best.row.answer, sourceDocument: best.row.sourceDocument, sourcePage: best.row.sourcePage, version: best.row.version, lastUpdated: best.row.updatedAt, demo: isDemoRecord_(best.row) });
    }
    var kb = runtimeRules_("KnowledgeBase", function(){ return true; });
    var bestKb = bestTextMatch_(kb, question, ["title","keywords","content"]);
    if (bestKb && bestKb.score > 0) {
      recordUsageInternal_(auth, "Q&A", bestKb.row.occupationCode || "", "", true, payload.startTime);
      return ok_({ answer: bestKb.row.content, sourceDocument: bestKb.row.sourceDocument, sourcePage: bestKb.row.sourcePage, version: bestKb.row.version, lastUpdated: bestKb.row.updatedAt, demo: isDemoRecord_(bestKb.row) });
    }
    recordUsageInternal_(auth, "Q&A", "", "", false, payload.startTime);
    return ok_({answer: "ไม่พบข้อมูลในฐานอ้างอิงที่ได้รับอนุมัติ กรุณาตรวจสอบคู่มือหรือสอบถามหน่วยงานที่เกี่ยวข้อง",sourceDocument: "-",sourcePage: "-",version: "-",lastUpdated: nowIso_(),notFound: true});
  } catch (error) { serverError_("askHomie", error); return fail_("ค้นคำตอบไม่สำเร็จ"); }
}

function uploadFile(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, ["Super Admin","Admin","Credit Editor"]);
    if (!auth.ok) return fail_(auth.message);
    var fileName = sanitizeFileName_(payload.fileName);
    var mimeType = cleanText_(payload.mimeType, 150);
    var base64 = String(payload.base64 || "");
    var targetFolder = cleanText_(payload.folder || "Uploads", 80);
    if (ALLOWED_UPLOAD_TYPES.indexOf(mimeType) < 0) return fail_("ประเภทไฟล์ไม่ได้รับอนุญาต");
    if (!base64) return fail_("ไม่พบข้อมูลไฟล์");
    var bytes = Utilities.base64Decode(base64);
    if (bytes.length > MAX_UPLOAD_BYTES) return fail_("ไฟล์มีขนาดเกิน 10 MB");
    var root = DriveApp.getFolderById(FOLDER_ID);
    var homie = getOrCreateChildFolder_(root, "Homie");
    var allowedFolders = ["Knowledge","Occupation","Product","Documents","Uploads"];
    if (allowedFolders.indexOf(targetFolder) < 0) targetFolder = "Uploads";
    var folder = getOrCreateChildFolder_(homie, targetFolder);
    var blob = Utilities.newBlob(bytes, mimeType, fileName);
    var file = folder.createFile(blob);
    writeFileRecord_(file, targetFolder, auth.user.username);
    writeLog_(auth.user.username, "UPLOAD", "Files", file.getId(), fileName, "", "");
    return ok_({ fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName(), fileType: mimeType, fileSize: bytes.length });
  } catch (error) { serverError_("uploadFile", error); return fail_("อัปโหลดไฟล์ไม่สำเร็จ"); }
}

function saveCaseLog(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var caseRef = cleanText_(payload.caseRef, 80);
    if (!/^CASE_[A-Za-z0-9_-]+$/.test(caseRef)) return fail_("caseRef ไม่ถูกต้อง");
    var row = {id: "CASELOG_" + Utilities.getUuid(),caseRef: caseRef,userId: auth.user.id,branch: auth.user.branch,occupationCode: cleanText_(payload.occupationCode, 50),purposeCode: cleanText_(payload.purposeCode, 50),module: cleanText_(payload.module, 100),resultSummary: cleanText_(payload.resultSummary, 1000),createdAt: nowIso_()};
    appendObject_("CaseLogs", row);
    return ok_(row);
  } catch (error) { serverError_("saveCaseLog", error); return fail_("บันทึก Case Log ไม่สำเร็จ"); }
}

function getUsageStats(payload) {
  try {
    var auth = requireSession_(payload && payload.token, ["Super Admin","Admin"]);
    if (!auth.ok) return fail_(auth.message);
    var usage = readSheetObjects_("UsageLogs");
    var feedback = readSheetObjects_("Feedback");
    var occupations = readSheetObjects_("Occupations");
    var purposes = readSheetObjects_("Purposes");
    var rulesCount = readSheetObjects_("IncomeRules").length + readSheetObjects_("ExpenseRules").length + readSheetObjects_("DSRRules").length + readSheetObjects_("ProductRules").length + readSheetObjects_("ChecklistRules").length + readSheetObjects_("RiskAlerts").length;
    return ok_({occupationsActive: countWhere_(occupations, function(r){ return isRuntimeMaster_(r); }),purposesActive: countWhere_(purposes, function(r){ return isRuntimeMaster_(r); }),rules: rulesCount,knowledgeArticles: runtimeRules_("KnowledgeBase", function(){ return true; }).length,checklistCount: countWhere_(usage, function(r){ return r.module === "Smart Checklist" && String(r.success).toUpperCase() === "TRUE"; }),incomeAssessments: countWhere_(usage, function(r){ return r.module === "Income Assessment"; }),users: readSheetObjects_("Users").length,sessions: readSheetObjects_("Sessions").length,feedbackScore: averageField_(feedback, "overallSatisfaction"),usageByOccupation: countBy_(usage, "occupationCode"),usageByBranch: countBy_(usage, "branch"),usageByModule: countBy_(usage, "module"),ruleErrors: countWhere_(readSheetObjects_("Logs"), function(r){ return r.action === "RULE_ERROR"; }),topQuestions: topQuestions_()});
  } catch (error) { serverError_("getUsageStats", error); return fail_("โหลดสถิติไม่สำเร็จ"); }
}

function recordUsage(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    recordUsageInternal_(auth, cleanText_(payload.module, 100), cleanText_(payload.occupationCode, 50), cleanText_(payload.purposeCode, 50), payload.success !== false, payload.startTime);
    return ok_(true);
  } catch (error) { serverError_("recordUsage", error); return fail_("บันทึก Usage ไม่สำเร็จ"); }
}

function saveFeedback(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var fields = ["easeOfUse","informationClarity","accuracyPerception","timeSaving","checklistUsefulness","overallSatisfaction"];
    var row = { id: "FB_" + Utilities.getUuid(), userId: auth.user.id, branch: auth.user.branch, createdAt: nowIso_() };
    var i;
    for (i = 0; i < fields.length; i++) {
      var n = numberOrNull_(payload[fields[i]]);
      if (n === null || n < 1 || n > 5) return fail_("คะแนน Feedback ต้องอยู่ระหว่าง 1–5");
      row[fields[i]] = n;
    }
    row.comment = cleanText_(payload.comment, 1500);
    appendObject_("Feedback", row);
    return ok_(row);
  } catch (error) { serverError_("saveFeedback", error); return fail_("บันทึก Feedback ไม่สำเร็จ"); }
}

function getFeedback(payload) { return getMaster_(payload, "Feedback", ["Super Admin","Admin"], false); }
function getUsageLogs(payload) { return getMaster_(payload, "UsageLogs", ["Super Admin","Admin"], false); }
function getLogs(payload) { return getMaster_(payload, "Logs", ["Super Admin","Admin"], false); }
function getSettings(payload) { return getMaster_(payload, "Settings", ["Super Admin","Admin"], false); }

function saveSettings(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, ["Super Admin","Admin"]);
    if (!auth.ok) return fail_(auth.message);
    var record = payload.record || {};
    var key = cleanText_(record.key, 100);
    if (!/^[A-Z0-9_]+$/.test(key)) return fail_("Settings key ไม่ถูกต้อง");
    var old = findOne_(readSheetObjects_("Settings"), "key", key);
    var row = { key: key, value: String(record.value || ""), updatedAt: nowIso_() };
    upsertObject_("Settings", "key", key, row);
    clearRuntimeCaches_();
    writeLog_(auth.user.username, old ? "UPDATE" : "CREATE", "Settings", key, "Settings changed", old ? JSON.stringify(old) : "", JSON.stringify(row));
    return ok_(row);
  } catch (error) { serverError_("saveSettings", error); return fail_("บันทึก Settings ไม่สำเร็จ"); }
}

function getUsers(payload) { return getMaster_(payload, "Users", ["Super Admin","Admin"], false, true); }

function saveUser(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, ["Super Admin","Admin"]);
    if (!auth.ok) return fail_(auth.message);
    var record = payload.record || {};
    var id = cleanText_(record.id, 80) || "USR_" + Utilities.getUuid();
    var username = cleanText_(record.username, 80).toLowerCase();
    if (!/^[a-z0-9._-]{3,80}$/.test(username)) return fail_("Username ไม่ถูกต้อง");
    var role = cleanText_(record.role, 50);
    if (ALLOWED_ROLES.indexOf(role) < 0) return fail_("Role ไม่ถูกต้อง");
    if (auth.user.role === "Admin" && role === "Super Admin") return fail_("Admin ไม่สามารถสร้างหรือแก้ Super Admin");
    var rows = readSheetObjects_("Users");
    var duplicate = null;
    var i;
    for (i = 0; i < rows.length; i++) if (rows[i].username === username && rows[i].id !== id) duplicate = rows[i];
    if (duplicate) return fail_("Username นี้มีอยู่แล้ว");
    var old = findOne_(rows, "id", id);
    var row = old ? clone_(old) : {};
    row.id = id; row.username = username; row.name = cleanText_(record.name, 200); row.branch = cleanText_(record.branch, 200); row.role = role;
    row.status = cleanStatus_(record.status || ACTIVE); row.mustChangePassword = String(record.mustChangePassword || (old ? old.mustChangePassword : "TRUE")).toUpperCase() === "TRUE" ? "TRUE" : "FALSE";
    row.createdAt = old ? old.createdAt : nowIso_(); row.updatedAt = nowIso_();
    if (!old || record.password) {
      var password = String(record.password || "");
      var valid = validatePassword_(password);
      if (!valid.ok) return fail_(valid.message);
      row.passwordSalt = randomToken_().substring(0, 40);
      row.passwordHash = hashPassword_(password, row.passwordSalt);
    }
    upsertObject_("Users", "id", id, row);
    writeLog_(auth.user.username, old ? "UPDATE" : "CREATE", "Users", id, "User saved", old ? JSON.stringify(redactUser_(old)) : "", JSON.stringify(redactUser_(row)));
    return ok_(publicUser_(row));
  } catch (error) { serverError_("saveUser", error); return fail_("บันทึกผู้ใช้ไม่สำเร็จ"); }
}

function deleteUser(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, ["Super Admin","Admin"]);
    if (!auth.ok) return fail_(auth.message);
    var id = cleanText_(payload.id, 80);
    if (id === auth.user.id) return fail_("ไม่สามารถลบบัญชีที่กำลังใช้งานอยู่");
    var target = findOne_(readSheetObjects_("Users"), "id", id);
    if (!target) return fail_("ไม่พบผู้ใช้");
    if (auth.user.role === "Admin" && target.role === "Super Admin") return fail_("Admin ไม่มีสิทธิ์ลบ Super Admin");
    deleteRowByKey_("Users", "id", id);
    revokeUserSessions_(id, "");
    writeLog_(auth.user.username, "DELETE", "Users", id, "User deleted", JSON.stringify(redactUser_(target)), "");
    return ok_(true);
  } catch (error) { serverError_("deleteUser", error); return fail_("ลบผู้ใช้ไม่สำเร็จ"); }
}

function createBackup(payload) {
  try {
    var auth = requireSession_(payload && payload.token, ["Super Admin","Admin"]);
    if (!auth.ok) return fail_(auth.message);
    var root = DriveApp.getFolderById(FOLDER_ID);
    var homie = getOrCreateChildFolder_(root, "Homie");
    var backup = getOrCreateChildFolder_(homie, "Backup");
    var name = "Homie_Backup_" + timestampForFile_();
    var copy = DriveApp.getFileById(SHEET_ID).makeCopy(name, backup);
    writeLog_(auth.user.username, "BACKUP", "System", copy.getId(), name, "", "");
    return ok_({ fileId: copy.getId(), fileUrl: copy.getUrl(), name: name });
  } catch (error) { serverError_("createBackup", error); return fail_("สำรองข้อมูลไม่สำเร็จ"); }
}

function createDemoData(payload) {
  try {
    var auth = requireSession_(payload && payload.token, ["Super Admin","Admin"]);
    if (!auth.ok) return fail_(auth.message);
    var now = nowIso_();
    setSettingInternal_("DEMO_MODE", "TRUE");
    var demoOcc = [
      { occupationId:"DEMO_OCC_001", occupationCode:"DEMO_OCC_001", occupationName:"ร้านอาหาร (TEST DATA)", occupationGroup:"DEMO", description:"ข้อมูลสำหรับทดสอบระบบเท่านั้น", requiredDocuments:'["หลักฐานรายรับ TEST DATA","หลักฐานค่าใช้จ่ายจริง TEST DATA"]', supportingDocuments:"[]", interviewQuestions:'["คำถามตัวอย่าง TEST DATA"]', incomeInputs:"[]", warning:"TEST DATA — ห้ามใช้พิจารณาสินเชื่อจริง", sourceDocument:"TEST DATA", sourcePage:"TEST DATA", version:"DEMO-1", effectiveDate:"2026-01-01", endDate:"2099-12-31", approvalStatus:APPROVED, status:ACTIVE, updatedBy:auth.user.username, updatedAt:now },
      { occupationId:"DEMO_OCC_002", occupationCode:"DEMO_OCC_002", occupationName:"ผู้ให้บริการรับส่ง / Rider (TEST DATA)", occupationGroup:"DEMO", description:"ข้อมูลสำหรับทดสอบระบบเท่านั้น", requiredDocuments:'["Statement รายได้ Platform TEST DATA","หลักฐานค่าใช้จ่ายจริง TEST DATA"]', supportingDocuments:"[]", interviewQuestions:'["คำถามตัวอย่าง TEST DATA"]', incomeInputs:"[]", warning:"TEST DATA — ห้ามใช้พิจารณาสินเชื่อจริง", sourceDocument:"TEST DATA", sourcePage:"TEST DATA", version:"DEMO-1", effectiveDate:"2026-01-01", endDate:"2099-12-31", approvalStatus:APPROVED, status:ACTIVE, updatedBy:auth.user.username, updatedAt:now }
    ];
    var i;
    for (i=0;i<demoOcc.length;i++) upsertObject_("Occupations","occupationId",demoOcc[i].occupationId, demoOcc[i]);
    var demoPurpose = { purposeId:"DEMO_PUR_001", purposeCode:"DEMO_PUR_001", purposeName:"ซื้อที่อยู่อาศัย (TEST DATA)", description:"TEST DATA", requiredCollateralDocuments:'["เอกสารหลักประกัน TEST DATA"]', optionalDocuments:"[]", specialConditions:"TEST DATA", warning:"TEST DATA", source:"TEST DATA", sourcePage:"TEST DATA", version:"DEMO-1", effectiveDate:"2026-01-01", endDate:"2099-12-31", approvalStatus:APPROVED, status:ACTIVE, updatedBy:auth.user.username, updatedAt:now };
    upsertObject_("Purposes","purposeId",demoPurpose.purposeId,demoPurpose);
    var demoRules = [
      { id:"DEMO_RULE_INC_001", occupationCode:"DEMO_OCC_001", ruleType:"INCOME", inputFields:'[{"key":"grossReceipts","label":"รายรับเฉลี่ยต่อเดือน (TEST DATA)","type":"number","required":true,"min":0},{"key":"actualExpense","label":"ค่าใช้จ่ายจริงต่อเดือน (TEST DATA)","type":"number","required":true,"min":0}]', formulaType:"DIRECT_FIELD", formulaConfig:'{"field":"grossReceipts"}', condition:"TEST DATA", source:"TEST DATA", sourcePage:"TEST DATA", version:"DEMO-1", effectiveDate:"2026-01-01", endDate:"2099-12-31", approvalStatus:APPROVED, status:ACTIVE, updatedBy:auth.user.username, updatedAt:now },
      { id:"DEMO_RULE_INC_002", occupationCode:"DEMO_OCC_002", ruleType:"INCOME", inputFields:'[{"key":"platformIncome","label":"รายได้เฉลี่ยจาก Platform ต่อเดือน (TEST DATA)","type":"number","required":true,"min":0},{"key":"actualExpense","label":"ค่าใช้จ่ายจริงต่อเดือน (TEST DATA)","type":"number","required":true,"min":0}]', formulaType:"DIRECT_FIELD", formulaConfig:'{"field":"platformIncome"}', condition:"TEST DATA", source:"TEST DATA", sourcePage:"TEST DATA", version:"DEMO-1", effectiveDate:"2026-01-01", endDate:"2099-12-31", approvalStatus:APPROVED, status:ACTIVE, updatedBy:auth.user.username, updatedAt:now }
    ];
    for(i=0;i<demoRules.length;i++) upsertObject_("IncomeRules","id",demoRules[i].id,demoRules[i]);
    var exp1 = { id:"DEMO_RULE_EXP_001", occupationCode:"DEMO_OCC_001", expenseType:"ACTUAL_FIELD", rate:"", formula:'{"field":"actualExpense"}', condition:"TEST DATA", source:"TEST DATA", sourcePage:"TEST DATA", version:"DEMO-1", effectiveDate:"2026-01-01", endDate:"2099-12-31", approvalStatus:APPROVED, status:ACTIVE, updatedBy:auth.user.username, updatedAt:now };
    var exp2 = clone_(exp1); exp2.id="DEMO_RULE_EXP_002"; exp2.occupationCode="DEMO_OCC_002";
    upsertObject_("ExpenseRules","id",exp1.id,exp1); upsertObject_("ExpenseRules","id",exp2.id,exp2);
    var dsr = { id:"DEMO_DSR_001", ruleName:"DSR TEST DATA", customerType:"ALL", productGroup:"ALL", incomeMin:"0", incomeMax:"", maxDSR:"50", condition:"TEST DATA ONLY", source:"TEST DATA", sourcePage:"TEST DATA", effectiveDate:"2026-01-01", endDate:"2099-12-31", version:"DEMO-1", approvalStatus:APPROVED, status:ACTIVE, updatedBy:auth.user.username, updatedAt:now };
    upsertObject_("DSRRules","id",dsr.id,dsr);
    var ck = [
      {id:"DEMO_CK_001",occupationCode:"DEMO_OCC_001",purposeCode:"DEMO_PUR_001",documentGroup:"Personal Documents",documentName:"บัตรประจำตัว (TEST DATA)",required:"TRUE",condition:"TEST DATA",customerText:"บัตรประจำตัว (TEST DATA)",staffNote:"TEST DATA",source:"TEST DATA",sourcePage:"TEST DATA",version:"DEMO-1",effectiveDate:"2026-01-01",endDate:"2099-12-31",approvalStatus:APPROVED,order:"1",status:ACTIVE,updatedBy:auth.user.username,updatedAt:now},
      {id:"DEMO_CK_002",occupationCode:"DEMO_OCC_001",purposeCode:"DEMO_PUR_001",documentGroup:"Income Documents",documentName:"หลักฐานรายรับร้านอาหาร (TEST DATA)",required:"TRUE",condition:"TEST DATA",customerText:"หลักฐานรายรับร้านอาหาร (TEST DATA)",staffNote:"TEST DATA",source:"TEST DATA",sourcePage:"TEST DATA",version:"DEMO-1",effectiveDate:"2026-01-01",endDate:"2099-12-31",approvalStatus:APPROVED,order:"2",status:ACTIVE,updatedBy:auth.user.username,updatedAt:now},
      {id:"DEMO_CK_003",occupationCode:"DEMO_OCC_002",purposeCode:"DEMO_PUR_001",documentGroup:"Income Documents",documentName:"Statement รายได้ Platform (TEST DATA)",required:"TRUE",condition:"TEST DATA",customerText:"Statement รายได้ Platform (TEST DATA)",staffNote:"TEST DATA",source:"TEST DATA",sourcePage:"TEST DATA",version:"DEMO-1",effectiveDate:"2026-01-01",endDate:"2099-12-31",approvalStatus:APPROVED,order:"1",status:ACTIVE,updatedBy:auth.user.username,updatedAt:now}
    ];
    for(i=0;i<ck.length;i++) upsertObject_("ChecklistRules","id",ck[i].id,ck[i]);
    var kb = { id:"DEMO_KB_001", category:"Demo", occupationCode:"DEMO_OCC_001", title:"ตัวอย่างการเตรียมเคสร้านอาหาร (TEST DATA)", content:"ใช้สำหรับทดสอบ Flow เท่านั้น ไม่ใช่หลักเกณฑ์สินเชื่อจริง", keywords:"ร้านอาหาร demo test", sourceDocument:"TEST DATA", sourcePage:"TEST DATA", version:"DEMO-1", effectiveDate:"2026-01-01", endDate:"2099-12-31", approvalStatus:APPROVED, status:ACTIVE, createdAt:now, updatedAt:now, updatedBy:auth.user.username };
    upsertObject_("KnowledgeBase","id",kb.id,kb);
    var qa = { id:"DEMO_QA_001", question:"ตัวอย่างข้อมูลร้านอาหาร", keywords:"ร้านอาหาร demo", answer:"ข้อมูลนี้เป็น TEST DATA สำหรับทดสอบระบบเท่านั้น", sourceDocument:"TEST DATA", sourcePage:"TEST DATA", version:"DEMO-1", effectiveDate:"2026-01-01", endDate:"2099-12-31", approvalStatus:APPROVED, status:ACTIVE, updatedAt:now, updatedBy:auth.user.username };
    upsertObject_("QALibrary","id",qa.id,qa);
    clearRuntimeCaches_();
    writeLog_(auth.user.username, "CREATE_DEMO", "System", "DEMO_", "Demo data created", "", "");
    return ok_({ created:true, demoMode:true, warning:"TEST DATA — ห้ามใช้พิจารณาสินเชื่อจริง" });
  } catch (error) { serverError_("createDemoData", error); return fail_("สร้าง Demo Data ไม่สำเร็จ"); }
}

function clearDemoData(payload) {
  try {
    var auth = requireSession_(payload && payload.token, ["Super Admin","Admin"]);
    if (!auth.ok) return fail_(auth.message);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheetName;
    var deleted = 0;
    for (sheetName in SHEET_SCHEMAS) {
      if (!SHEET_SCHEMAS.hasOwnProperty(sheetName)) continue;
      if (sheetName === "Logs" || sheetName === "Sessions") continue;
      var sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) continue;
      var values = sh.getDataRange().getValues();
      var headers = values[0];
      var idIndexes = [];
      var h;
      for (h=0;h<headers.length;h++) if (headers[h] === "id" || headers[h] === "occupationId" || headers[h] === "purposeId" || headers[h] === "productId" || headers[h] === "caseRef" || headers[h] === "key") idIndexes.push(h);
      var r;
      for (r=values.length-1;r>=1;r--) {
        var demo = false;
        var x;
        for (x=0;x<idIndexes.length;x++) if (String(values[r][idIndexes[x]] || "").indexOf("DEMO_") === 0) demo = true;
        if (demo) { sh.deleteRow(r+1); deleted++; }
      }
    }
    setSettingInternal_("DEMO_MODE", "FALSE");
    clearRuntimeCaches_();
    writeLog_(auth.user.username, "CLEAR_DEMO", "System", "DEMO_", "Demo data cleared: " + deleted, "", "");
    return ok_({ deleted:deleted, demoMode:false });
  } catch (error) { serverError_("clearDemoData", error); return fail_("ล้าง Demo Data ไม่สำเร็จ"); }
}

function importMasterData(payload) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, ["Super Admin","Admin"]);
    if (!auth.ok) return fail_(auth.message);
    var sheetName = cleanText_(payload.sheetName, 80);
    if (!SHEET_SCHEMAS[sheetName]) return fail_("ไม่รองรับ Sheet นี้");
    if (["Users","Sessions","UsageLogs","CaseLogs","Feedback","Files","Logs"].indexOf(sheetName) >= 0) return fail_("Sheet นี้ไม่อนุญาตให้นำเข้าผ่าน Master Import");
    var rows = safeArray_(payload.rows);
    if (!rows.length) return fail_("ไม่พบข้อมูลนำเข้า");
    if (rows.length > 2000) return fail_("นำเข้าได้สูงสุดครั้งละ 2,000 รายการ");
    var keyField = primaryKeyForSheet_(sheetName);
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var i;
      for (i=0;i<rows.length;i++) {
        var record = sanitizeRecordForSheet_(sheetName, rows[i]);
        validateImportedRule_(sheetName, record);
        var key = record[keyField];
        if (!key) throw new Error("Row " + (i+1) + " ไม่มี " + keyField);
        upsertObject_(sheetName, keyField, key, record);
      }
    } finally { lock.releaseLock(); }
    clearRuntimeCaches_();
    writeLog_(auth.user.username, "IMPORT", sheetName, "", "Imported " + rows.length + " rows", "", "");
    return ok_({ sheetName:sheetName, imported:rows.length });
  } catch (error) { serverError_("importMasterData", error); return fail_(friendlyError_(error)); }
}

function safeArray_(value) { return Array.isArray(value) ? value : []; }
function safeJson_(value, fallback) {
  if (value === null || value === undefined || value === "") return fallback;
  if (typeof value === "object") {
    if (Array.isArray(fallback)) return Array.isArray(value) ? value : fallback;
    if (fallback && typeof fallback === "object") return !Array.isArray(value) ? value : fallback;
    return value;
  }
  try {
    var parsed = JSON.parse(value);
    if (Array.isArray(fallback)) return Array.isArray(parsed) ? parsed : fallback;
    if (fallback && typeof fallback === "object") return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : fallback;
    return parsed === null || parsed === undefined ? fallback : parsed;
  } catch (error) { return fallback; }
}

function clearRuntimeCaches_() {
  var cache = CacheService.getScriptCache();
  var keys = ["Settings","Occupations","Purposes","DSRRules","ProductRules","ChecklistRules","KnowledgeBase","DashboardStats","IncomeRules","ExpenseRules","Products","RiskAlerts","QALibrary"];
  var i; for(i=0;i<keys.length;i++) cache.remove("HOMIE_" + keys[i]);
}

function getMaster_(payload, sheetName, roles, runtimeForStaff, redact) {
  try {
    var auth = requireSession_(payload && payload.token, roles);
    if (!auth.ok) return fail_(auth.message);
    var rows = readSheetObjects_(sheetName);
    if (runtimeForStaff && !isAdminRole_(auth.user.role)) rows = filterBy_(rows, function(r){ return isRuntimeMaster_(r); });
    if (redact && sheetName === "Users") rows = mapBy_(rows, function(r){ return publicUser_(r); });
    return ok_(rows);
  } catch (error) { serverError_("getMaster_ " + sheetName, error); return fail_("โหลดข้อมูลไม่สำเร็จ"); }
}

function saveMaster_(payload, sheetName, keyField, prefix, roles, critical) {
  var lock = null;
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, roles);
    if (!auth.ok) return fail_(auth.message);
    var record = sanitizeRecordForSheet_(sheetName, payload.record || {});
    var key = cleanText_(record[keyField], 120);
    if (!key) { key = prefix + Utilities.getUuid(); record[keyField] = key; }
    validateRecordId_(key);
    if (critical) validateCriticalRuleForSave_(record);
    record.updatedBy = auth.user.username;
    record.updatedAt = nowIso_();
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    var old = findOne_(readSheetObjects_(sheetName), keyField, key);
    upsertObject_(sheetName, keyField, key, record);
    lock.releaseLock(); lock = null;
    clearRuntimeCaches_();
    writeLog_(auth.user.username, old ? "UPDATE" : "CREATE", sheetName, key, critical ? "Critical master saved" : "Master saved", old ? JSON.stringify(old) : "", JSON.stringify(record));
    return ok_(record);
  } catch (error) {
    try { if (lock) lock.releaseLock(); } catch(ignore) {}
    serverError_("saveMaster_ " + sheetName, error);
    return fail_(friendlyError_(error));
  }
}

function deleteMaster_(payload, sheetName, keyField, roles) {
  try {
    payload = payload || {};
    var auth = requireSession_(payload.token, roles);
    if (!auth.ok) return fail_(auth.message);
    var id = cleanText_(payload.id, 120);
    validateRecordId_(id);
    var old = findOne_(readSheetObjects_(sheetName), keyField, id);
    if (!old) return fail_("ไม่พบข้อมูลที่ต้องการลบ");
    deleteRowByKey_(sheetName, keyField, id);
    clearRuntimeCaches_();
    writeLog_(auth.user.username, "DELETE", sheetName, id, "Master deleted", JSON.stringify(old), "");
    return ok_(true);
  } catch (error) { serverError_("deleteMaster_ " + sheetName, error); return fail_(friendlyError_(error)); }
}

function getKnowledgeItem_(payload) {
  try {
    var auth = requireSession_(payload && payload.token, null);
    if (!auth.ok) return fail_(auth.message);
    var id = cleanText_(payload.id, 120);
    var row = findOne_(readSheetObjects_("KnowledgeBase"), "id", id);
    if (!row) return fail_("ไม่พบข้อมูล");
    if (!isAdminRole_(auth.user.role) && !isRuntimeRule_(row)) return fail_(approvedMissingMessage_());
    return ok_(row);
  } catch (error) { serverError_("getKnowledgeById", error); return fail_("โหลดบทความไม่สำเร็จ"); }
}

function validateConfig_() {
  if (!SHEET_ID || SHEET_ID === "XXX") throw new Error("กรุณาตั้งค่า SHEET_ID ก่อน Run setupSystem()");
  if (!FOLDER_ID || FOLDER_ID === "XXX") throw new Error("กรุณาตั้งค่า FOLDER_ID ก่อน Run setupSystem()");
  if (!ADMIN_PASSWORD || ADMIN_PASSWORD === "CHANGE_ME_BEFORE_SETUP" || ADMIN_PASSWORD === "1234" || ADMIN_PASSWORD.length < 10) throw new Error("กรุณาเปลี่ยน ADMIN_PASSWORD เป็นรหัสผ่าน Bootstrap ที่เดายากอย่างน้อย 10 ตัวอักษรก่อน setup");
  SpreadsheetApp.openById(SHEET_ID);
  DriveApp.getFolderById(FOLDER_ID);
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  var created = false;
  if (!sh) { sh = ss.insertSheet(name); created = true; }
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    var current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    var i; for(i=0;i<headers.length;i++) if (String(current[i] || "") !== headers[i]) throw new Error("Schema mismatch: " + name + " column " + (i+1) + " ต้องเป็น " + headers[i]);
  }
  return created;
}

function ensureDriveFolders_(root) {
  var names = ["Knowledge","Occupation","Product","Documents","Exported_Checklists","Backup","Uploads"];
  var homie = getOrCreateChildFolder_(root, "Homie");
  var result = ["Homie"];
  var i; for(i=0;i<names.length;i++){ getOrCreateChildFolder_(homie, names[i]); result.push("Homie/"+names[i]); }
  return result;
}

function seedSettings_() {
  var defaults = {SYSTEM_NAME:"Homie Credit Assistant",ORGANIZATION:"",REGION:"",DISCLAIMER:"ผลการประเมินเป็นข้อมูลเบื้องต้นสำหรับช่วยพนักงานเตรียมเคสเท่านั้น การพิจารณาสินเชื่อต้องเป็นไปตามหลักเกณฑ์ คู่มือ และอำนาจอนุมัติของธนาคารที่มีผลบังคับใช้",MANUAL_NAME:"",MANUAL_VERSION:"",MANUAL_EFFECTIVE_DATE:"",LOAN_ESTIMATE_CONFIG:"{}",THEME:"light",PRIMARY_COLOR:"#0B1F3A",SECONDARY_COLOR:"#1E5EFF",ACCENT_COLOR:"#F59E0B",FOOTER_TEXT:"Homie Credit Assistant — Staff Credit Assistant",CONTACT_SUPPORT:"",DEMO_MODE:"FALSE"};
  var k; for(k in defaults) if(defaults.hasOwnProperty(k) && !findOne_(readSheetObjects_("Settings"),"key",k)) setSettingInternal_(k,defaults[k]);
}

function seedOccupationCatalog_() {
  var rows = readSheetObjects_("Occupations");
  var i;
  for(i=0;i<OCCUPATION_CATALOG.length;i++) if (!findOne_(rows,"occupationCode",OCCUPATION_CATALOG[i][0])) appendObject_("Occupations", {occupationId:"CAT_"+OCCUPATION_CATALOG[i][0], occupationCode:OCCUPATION_CATALOG[i][0], occupationName:OCCUPATION_CATALOG[i][1], occupationGroup:"",description:"Catalog only — ต้องนำเข้า Homie V7 Master ก่อนใช้งาน Runtime", characteristics:"", requiredDocuments:"[]", supportingDocuments:"[]", interviewQuestions:"[]", incomeInputs:"[]", incomeMethod:"", incomeFormula:"", expenseMethod:"", expenseRate:"", expenseConditions:"", warning:"ยังไม่อนุมัติสำหรับ Runtime", riskPoints:"[]", calculationNotes:"", sourceDocument:"", sourcePage:"", version:"", effectiveDate:"", endDate:"", approvalStatus:"REVIEW_REQUIRED", status:"INACTIVE", updatedBy:"system", updatedAt:nowIso_()});
}

function seedAdmin_() {
  var users = readSheetObjects_("Users");
  if (findOne_(users,"username","admin")) return;
  var salt = randomToken_().substring(0,40);
  appendObject_("Users", {id:"USR_ADMIN", username:"admin", passwordHash:hashPassword_(ADMIN_PASSWORD,salt), passwordSalt:salt, name:"System Administrator", branch:"", role:"Super Admin", status:ACTIVE, mustChangePassword:"TRUE", createdAt:nowIso_(), updatedAt:nowIso_()});
}

function validateSchema_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var result = {};
  var name;
  for(name in SHEET_SCHEMAS) {
    if(!SHEET_SCHEMAS.hasOwnProperty(name)) continue;
    var sh=ss.getSheetByName(name); result[name]=!!sh;
    if(!sh) throw new Error("Missing sheet: "+name);
    var headers=sh.getRange(1,1,1,SHEET_SCHEMAS[name].length).getValues()[0];
    var i; for(i=0;i<SHEET_SCHEMAS[name].length;i++) if(headers[i]!==SHEET_SCHEMAS[name][i]) throw new Error("Schema mismatch: "+name+"."+SHEET_SCHEMAS[name][i]);
  }
  return result;
}

function readSheetObjects_(sheetName) {
  var cacheable = ["Settings","Occupations","Purposes","IncomeRules","ExpenseRules","DSRRules","Products","ProductRules","ChecklistRules","RiskAlerts","KnowledgeBase","QALibrary"].indexOf(sheetName) >= 0;
  var cache = CacheService.getScriptCache();
  var cacheKey = "HOMIE_" + sheetName;
  if (cacheable) {
    var cached = cache.get(cacheKey);
    if (cached) { var parsed = safeJson_(cached, []); if (Array.isArray(parsed)) return parsed; }
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var rows = [];
  var r,c;
  for(r=1;r<values.length;r++) {
    var empty = true; var obj={};
    for(c=0;c<headers.length;c++) {
      var v=values[r][c]; if(v!=="" && v!==null) empty=false;
      if(v instanceof Date) v=Utilities.formatDate(v, Session.getScriptTimeZone() || "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm:ssXXX");
      obj[headers[c]]=v;
    }
    if(!empty) rows.push(obj);
  }
  if(cacheable) { try { var text=JSON.stringify(rows); if(text.length<90000) cache.put(cacheKey,text,300); } catch(ignore) {} }
  return rows;
}

function appendObject_(sheetName,obj) {
  var headers=SHEET_SCHEMAS[sheetName]; if(!headers) throw new Error("Unknown sheet "+sheetName);
  var sh=SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
  var row=[]; var i; for(i=0;i<headers.length;i++) row.push(sanitizeSheetCell_(obj[headers[i]]));
  sh.appendRow(row); CacheService.getScriptCache().remove("HOMIE_"+sheetName);
}
function upsertObject_(sheetName,keyField,key,obj) { var updated=updateRowByKey_(sheetName,keyField,key,obj); if(!updated) appendObject_(sheetName,obj); }
function updateRowByKey_(sheetName,keyField,key,obj) {
  var headers=SHEET_SCHEMAS[sheetName]; var sh=SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
  if(!sh || sh.getLastRow()<2) return false;
  var keyIndex=headers.indexOf(keyField); if(keyIndex<0) throw new Error("Missing key field "+keyField);
  var values=sh.getRange(2,keyIndex+1,sh.getLastRow()-1,1).getValues(); var i;
  for(i=0;i<values.length;i++) if(String(values[i][0])===String(key)) {
    var row=[]; var c; for(c=0;c<headers.length;c++) row.push(sanitizeSheetCell_(obj[headers[c]]));
    sh.getRange(i+2,1,1,headers.length).setValues([row]); CacheService.getScriptCache().remove("HOMIE_"+sheetName); return true;
  }
  return false;
}
function deleteRowByKey_(sheetName,keyField,key) {
  var headers=SHEET_SCHEMAS[sheetName]; var sh=SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName); if(!sh || sh.getLastRow()<2) return false;
  var idx=headers.indexOf(keyField); var values=sh.getRange(2,idx+1,sh.getLastRow()-1,1).getValues(); var i;
  for(i=values.length-1;i>=0;i--) if(String(values[i][0])===String(key)){ sh.deleteRow(i+2); CacheService.getScriptCache().remove("HOMIE_"+sheetName); return true; }
  return false;
}
function sanitizeSheetCell_(value) { if(value===null || value===undefined) return ""; if(typeof value === "object") value = JSON.stringify(value); var text=String(value); if(/^[=+\-@]/.test(text)) return "'"+text; return text; }
function sanitizeRecordForSheet_(sheetName,record) { var out={}; var headers=SHEET_SCHEMAS[sheetName]; var i; for(i=0;i<headers.length;i++) {var key=headers[i]; var value=record[key]; if(value===undefined || value===null) value=""; if(typeof value === "object") value=JSON.stringify(value); out[key]=cleanCellInput_(value,10000);} return out; }
function cleanCellInput_(value,maxLen) { var s=String(value===undefined||value===null?"":value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,""); if(s.length>maxLen) s=s.substring(0,maxLen); return s; }

function requireSession_(token,roles) {
  token=String(token||""); if(!token) return {ok:false,message:"กรุณาเข้าสู่ระบบใหม่"};
  var hash=sha256_(token); var sessions=readSheetObjects_("Sessions"); var session=findOne_(sessions,"tokenHash",hash);
  if(!session || String(session.revoked).toUpperCase()==="TRUE") return {ok:false,message:"Session หมดอายุหรือถูกยกเลิก กรุณาเข้าสู่ระบบใหม่"};
  var expires=new Date(session.expiresAt); if(isNaN(expires.getTime()) || expires.getTime()<=new Date().getTime()) return {ok:false,message:"Session หมดอายุ กรุณาเข้าสู่ระบบใหม่"};
  var user=getUserById_(session.userId); if(!user || String(user.status).toUpperCase()!==ACTIVE) return {ok:false,message:"บัญชีผู้ใช้ไม่พร้อมใช้งาน"};
  if(roles && roles.length && roles.indexOf(user.role)<0 && user.role!=="Super Admin") return {ok:false,message:"คุณไม่มีสิทธิ์ใช้งานรายการนี้"};
  session.lastAccess=nowIso_(); updateRowByKey_("Sessions","tokenHash",hash,session);
  return {ok:true,user:user,session:session};
}
function createSession_(user) { var token=randomToken_(); var expires=new Date(new Date().getTime()+SESSION_HOURS*60*60*1000).toISOString(); appendObject_("Sessions",{tokenHash:sha256_(token),userId:user.id,expiresAt:expires,revoked:"FALSE",createdAt:nowIso_(),lastAccess:nowIso_()}); return {token:token,expiresAt:expires}; }
function revokeUserSessions_(userId,currentToken) { var rows=readSheetObjects_("Sessions"); var currentHash=currentToken?sha256_(currentToken):""; var i; for(i=0;i<rows.length;i++) if(rows[i].userId===userId && rows[i].tokenHash!==currentHash && String(rows[i].revoked).toUpperCase()!=="TRUE") { rows[i].revoked="TRUE"; updateRowByKey_("Sessions","tokenHash",rows[i].tokenHash,rows[i]); } }
function countUserSessions_(userId) { return countWhere_(readSheetObjects_("Sessions"),function(r){return r.userId===userId;}); }
function getUserById_(id){ return findOne_(readSheetObjects_("Users"),"id",id); }
function publicUser_(u){ return {id:u.id,username:u.username,name:u.name,branch:u.branch,role:u.role,status:u.status,mustChangePassword:u.mustChangePassword}; }
function redactUser_(u){ var x=clone_(u); x.passwordHash="[REDACTED]"; x.passwordSalt="[REDACTED]"; return x; }
function isAdminRole_(role){ return ["Super Admin","Admin","Credit Editor"].indexOf(role)>=0; }
function hashPassword_(password,salt) { var value=String(password)+"|"+String(salt); var i; for(i=0;i<PASSWORD_ROUNDS;i++) value=sha256_(value+"|"+salt+"|"+i); return value; }
function sha256_(text){ var bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(text),Utilities.Charset.UTF_8); var out=""; var i; for(i=0;i<bytes.length;i++){ var v=bytes[i]; if(v<0)v+=256; var h=v.toString(16); if(h.length<2)h="0"+h; out+=h; } return out; }
function safeEqual_(a,b){ a=String(a||""); b=String(b||""); if(a.length!==b.length)return false; var diff=0; var i; for(i=0;i<a.length;i++)diff|=a.charCodeAt(i)^b.charCodeAt(i); return diff===0; }
function randomToken_(){ return Utilities.getUuid().replace(/-/g,"")+Utilities.getUuid().replace(/-/g,"")+Utilities.getUuid().replace(/-/g,"")+new Date().getTime(); }
function validatePassword_(p){ if(!p || p.length<10)return {ok:false,message:"รหัสผ่านใหม่ต้องมีอย่างน้อย 10 ตัวอักษร"}; if(!/[A-Za-z]/.test(p)||!/[0-9]/.test(p))return {ok:false,message:"รหัสผ่านต้องมีทั้งตัวอักษรและตัวเลข"}; return {ok:true}; }

function runtimeRules_(sheetName,predicate) { var rows=readSheetObjects_(sheetName); var out=[]; var i; for(i=0;i<rows.length;i++) if(isRuntimeRule_(rows[i]) && (!predicate || predicate(rows[i]))) out.push(rows[i]); return out; }
function isRuntimeRule_(r) { if(!r) return false; var demo=isDemoRecord_(r); if(demo && String(getSettingValue_("DEMO_MODE")).toUpperCase()!=="TRUE") return false; if(String(r.status).toUpperCase()!==ACTIVE) return false; if(String(r.approvalStatus).toUpperCase()!==APPROVED) return false; var source=r.source || r.sourceDocument; if(!source || !r.version) return false; if(!isDateEffective_(r.effectiveDate,r.endDate || r.expiryDate)) return false; return true; }
function isRuntimeMaster_(r) { return isRuntimeRule_(r); }
function isDateEffective_(start,end){ var now=new Date(); if(start){var s=new Date(start); if(isNaN(s.getTime())||s.getTime()>now.getTime())return false;} if(end){var e=new Date(end); if(!isNaN(e.getTime())&&e.getTime()<now.getTime())return false;} return true; }
function isDemoRecord_(r){ var fields=[r.id,r.occupationId,r.purposeId,r.productId,r.occupationCode,r.purposeCode]; var i; for(i=0;i<fields.length;i++) if(String(fields[i]||"").indexOf("DEMO_")===0)return true; return false; }
function findRuntimeOccupation_(code){ var rows=runtimeRules_("Occupations",function(r){return r.occupationCode===code;}); if(rows.length!==1)return null; return rows[0]; }
function findRuntimePurpose_(code){ var rows=runtimeRules_("Purposes",function(r){return r.purposeCode===code;}); if(rows.length!==1)return null; return rows[0]; }
function approvedMissingMessage_(){ return "ไม่พบข้อมูลในฐานอ้างอิงที่ได้รับอนุมัติ กรุณาตรวจสอบคู่มือหรือสอบถามหน่วยงานที่เกี่ยวข้อง"; }
function validateCriticalRuleForSave_(r) { if(String(r.status).toUpperCase()===ACTIVE) {var source=r.source||r.sourceDocument; if(!source) throw new Error("Rule Active ต้องมี Source"); if(!r.version) throw new Error("Rule Active ต้องมี Version"); if(!r.effectiveDate) throw new Error("Rule Active ต้องมี Effective Date"); if(String(r.approvalStatus).toUpperCase()!==APPROVED) throw new Error("Rule Active ต้องมี Approval Status = APPROVED");} }
function validateImportedRule_(sheetName,r){ if(["IncomeRules","ExpenseRules","DSRRules","Products","ProductRules","ChecklistRules","RiskAlerts","KnowledgeBase","QALibrary"].indexOf(sheetName)>=0) validateCriticalRuleForSave_(r); }
function validateRecordId_(id){ if(!/^[A-Za-z0-9_.-]{1,120}$/.test(String(id||""))) throw new Error("Record ID ไม่ถูกต้อง"); }
function cleanStatus_(s){ s=String(s||"").toUpperCase(); return ["ACTIVE","INACTIVE","DRAFT","ARCHIVED"].indexOf(s)>=0?s:"INACTIVE"; }

function validateDynamicInputs_(schema,inputs) { schema=safeArray_(schema); var i; for(i=0;i<schema.length;i++) {var f=schema[i]||{}; var key=String(f.key||""); if(!key) return {ok:false,message:"Income Rule มี inputFields ไม่สมบูรณ์"}; var v=inputs[key]; if(f.required && (v===""||v===null||v===undefined))return {ok:false,message:"กรุณากรอก "+(f.label||key)}; if(f.type==="number" && v!=="" && v!==null && v!==undefined){ var n=Number(v); if(!isFinite(n))return {ok:false,message:(f.label||key)+" ต้องเป็นตัวเลข"}; if(f.min!==undefined && n<Number(f.min))return {ok:false,message:(f.label||key)+" ต้องไม่น้อยกว่า "+f.min}; if(f.max!==undefined && n>Number(f.max))return {ok:false,message:(f.label||key)+" ต้องไม่เกิน "+f.max}; }} return {ok:true}; }
function computeFormula_(rule,inputs) { var type=String(rule.formulaType||"").toUpperCase(); var cfg=safeJson_(rule.formulaConfig,{}); var value=0; var detail=[]; var i; if(type==="DIRECT_FIELD") { var field=cfg.field; var n=numberOrNull_(inputs[field]); if(n===null||n<0)return {ok:false,message:"ข้อมูลสำหรับคำนวณรายได้ไม่ครบ"}; value=n; detail.push(field+" = "+n); } else if(type==="SUM_FIELDS") { var fields=safeArray_(cfg.fields); if(!fields.length)return {ok:false,message:"formulaConfig ไม่มี fields"}; for(i=0;i<fields.length;i++){var sn=numberOrNull_(inputs[fields[i]]); if(sn===null)return {ok:false,message:"ข้อมูล "+fields[i]+" ไม่ครบ"}; value+=sn; detail.push(fields[i]+" = "+sn);} } else if(type==="PRODUCT_FIELDS") { var pfields=safeArray_(cfg.fields); if(!pfields.length)return {ok:false,message:"formulaConfig ไม่มี fields"}; value=1; for(i=0;i<pfields.length;i++){var pn=numberOrNull_(inputs[pfields[i]]); if(pn===null)return {ok:false,message:"ข้อมูล "+pfields[i]+" ไม่ครบ"}; value*=pn; detail.push(pfields[i]+" = "+pn);} } else if(type==="AVERAGE_FIELDS") { var af=safeArray_(cfg.fields); if(!af.length)return {ok:false,message:"formulaConfig ไม่มี fields"}; var sum=0; for(i=0;i<af.length;i++){var an=numberOrNull_(inputs[af[i]]); if(an===null)return {ok:false,message:"ข้อมูล "+af[i]+" ไม่ครบ"}; sum+=an; detail.push(af[i]+" = "+an);} value=sum/af.length; } else return {ok:false,message:"ไม่รองรับ formulaType: "+type+" กรุณาตรวจสอบ Rule Master"}; if(!isFinite(value)||value<0)return {ok:false,message:"ผลคำนวณรายได้ไม่ถูกต้อง"}; detail.push("Gross Income = "+round2_(value)); return {ok:true,value:value,detail:detail}; }
function computeExpense_(rule,inputs,gross) { if(!rule) return {ok:false,message:"ไม่สามารถประเมินได้ เนื่องจากข้อมูล Rule ค่าใช้จ่ายไม่สมบูรณ์"}; var type=String(rule.expenseType||"").toUpperCase(); var value=0; var detail=[]; if(type==="ACTUAL_FIELD") { var cfg=safeJson_(rule.formula,{}); var field=cfg.field; var n=numberOrNull_(inputs[field]); if(n===null||n<0)return {ok:false,message:"กรุณากรอกค่าใช้จ่ายจริงตาม Rule"}; value=n; detail.push("Actual expense "+field+" = "+n); } else if(type==="PERCENT_OF_GROSS") { var rate=numberOrNull_(rule.rate); if(rate===null||rate<0||rate>100)return {ok:false,message:"Expense Rule ไม่มี rate ที่ถูกต้อง"}; value=gross*rate/100; detail.push("Expense = Gross x "+rate+"%"); } else if(type==="FIXED") { var fixed=numberOrNull_(rule.rate); if(fixed===null||fixed<0)return {ok:false,message:"Expense Rule ไม่มีค่าที่ถูกต้อง"}; value=fixed; detail.push("Fixed expense = "+fixed); } else return {ok:false,message:"ไม่รองรับ expenseType: "+type}; return {ok:true,value:value,detail:detail}; }
function evaluateCondition_(actual,operator,expected) { operator=String(operator||"EQ").toUpperCase(); var exp=safeJson_(expected,expected); if(operator==="EQ") return String(actual)===String(exp); if(operator==="NEQ") return String(actual)!==String(exp); if(operator==="GTE") return numberCompare_(actual,exp,function(a,b){return a>=b;}); if(operator==="LTE") return numberCompare_(actual,exp,function(a,b){return a<=b;}); if(operator==="GT") return numberCompare_(actual,exp,function(a,b){return a>b;}); if(operator==="LT") return numberCompare_(actual,exp,function(a,b){return a<b;}); if(operator==="IN") { var arr=Array.isArray(exp)?exp:String(exp).split(","); var i; for(i=0;i<arr.length;i++)if(String(arr[i]).trim()===String(actual))return true; return false; } if(operator==="CONTAINS") return String(actual||"").toLowerCase().indexOf(String(exp||"").toLowerCase())>=0; if(operator==="BETWEEN") { var range=Array.isArray(exp)?exp:String(exp).split(","); if(range.length<2)return false; var n=numberOrNull_(actual),a=numberOrNull_(range[0]),b=numberOrNull_(range[1]); return n!==null&&a!==null&&b!==null&&n>=a&&n<=b; } if(operator==="EMPTY") return actual===""||actual===null||actual===undefined; if(operator==="NOT_EMPTY") return !(actual===""||actual===null||actual===undefined); return false; }
function numberCompare_(a,b,fn){ a=numberOrNull_(a); b=numberOrNull_(b); return a!==null&&b!==null?fn(a,b):false; }
function getApprovedLoanEstimateConfig_() { var raw=getSettingValue_("LOAN_ESTIMATE_CONFIG"); var cfg=safeJson_(raw,{}); if(!cfg || String(cfg.approvalStatus).toUpperCase()!==APPROVED || !cfg.source || !cfg.version || !cfg.effectiveDate || !isDateEffective_(cfg.effectiveDate,cfg.endDate)) return {ok:false,message:"ไม่พบ Loan Estimate Rule/Settings ที่ได้รับอนุมัติ"}; return {ok:true,data:cfg}; }
function sourceView_(r){ if(!r)return null; return {source:r.source||r.sourceDocument||"",sourceDocument:r.sourceDocument||r.source||"",sourcePage:r.sourcePage||"",version:r.version||"",effectiveDate:r.effectiveDate||"",approvalStatus:r.approvalStatus||"",status:r.status||"",demo:isDemoRecord_(r)}; }
function collectWarnings_(occ,inc,exp){ var out=[]; if(occ.warning)out.push(occ.warning); if(inc.condition)out.push(inc.condition); if(exp&&exp.condition)out.push(exp.condition); return out; }
function isApprovalQuestion_(q){ q=String(q||"").toLowerCase(); var keys=["ผ่านไหม","อนุมัติได้","ควรอนุมัติ","reject","approved","อนุมัติไหม","ผ่านหรือไม่"]; var i; for(i=0;i<keys.length;i++)if(q.indexOf(keys[i])>=0)return true; return false; }
function bestTextMatch_(rows,question,fields){ var tokens=tokenize_(question); var best=null; var i,j,k; for(i=0;i<rows.length;i++){var text=""; for(j=0;j<fields.length;j++)text+=" "+String(rows[i][fields[j]]||"").toLowerCase(); var score=0; for(k=0;k<tokens.length;k++)if(tokens[k].length>1 && text.indexOf(tokens[k])>=0)score++; if(score>0 && (!best||score>best.score))best={row:rows[i],score:score};} return best; }
function tokenize_(text){ return String(text||"").toLowerCase().replace(/[^0-9a-zA-Zก-๙]+/g," ").split(/\s+/); }
function recordUsageInternal_(auth,module,occ,pur,success,startTime){ var start=startTime?new Date(startTime):new Date(); if(isNaN(start.getTime()))start=new Date(); var end=new Date(); appendObject_("UsageLogs",{id:"USE_"+Utilities.getUuid(),sessionId:auth.session.tokenHash.substring(0,16),userId:auth.user.id,branch:auth.user.branch,module:module,occupationCode:occ,purposeCode:pur,startTime:start.toISOString(),endTime:end.toISOString(),durationSeconds:Math.max(0,Math.round((end.getTime()-start.getTime())/1000)),success:success?"TRUE":"FALSE",createdAt:nowIso_()}); }
function topQuestions_(){ var logs=readSheetObjects_("UsageLogs"); return countBy_(filterBy_(logs,function(r){return r.module==="Q&A";}),"occupationCode"); }
function writeFileRecord_(file,folder,username){ var key=""; try{key=file.getResourceKey?file.getResourceKey()||"":"";}catch(ignore){} appendObject_("Files",{id:"FILE_"+Utilities.getUuid(),fileId:file.getId(),fileUrl:file.getUrl(),fileName:file.getName(),fileType:file.getMimeType(),fileSize:file.getSize(),folder:folder,resourceKey:key,createdAt:nowIso_(),uploadedBy:username}); }
function getOrCreateChildFolder_(parent,name){ var it=parent.getFoldersByName(name); return it.hasNext()?it.next():parent.createFolder(name); }
function sanitizeFileName_(name){ var s=cleanText_(name,180).replace(/[\\/:*?"<>|]/g,"_"); if(!s)throw new Error("ชื่อไฟล์ไม่ถูกต้อง"); return s; }
function getSettingValue_(key){ var rows=readSheetObjects_("Settings"); var row=findOne_(rows,"key",key); return row?row.value:""; }
function setSettingInternal_(key,value){ upsertObject_("Settings","key",key,{key:key,value:String(value),updatedAt:nowIso_()}); }
function primaryKeyForSheet_(name){ var map={Settings:"key",Occupations:"occupationId",Purposes:"purposeId",IncomeRules:"id",ExpenseRules:"id",DSRRules:"id",Products:"productId",ProductRules:"id",ChecklistRules:"id",RiskAlerts:"id",KnowledgeBase:"id",QALibrary:"id"}; return map[name]||"id"; }
function writeLog_(username,action,module,recordId,detail,oldValue,newValue){ try{ appendObject_("Logs",{id:"LOG_"+Utilities.getUuid(),username:cleanText_(username,100),action:cleanText_(action,80),module:cleanText_(module,100),recordId:cleanText_(recordId,150),detail:cleanText_(detail,2000),oldValue:cleanText_(oldValue,9000),newValue:cleanText_(newValue,9000),timestamp:nowIso_()}); }catch(ignore){} }
function serverError_(where,error){ try{ console.error(where+": "+(error&&error.stack?error.stack:error)); writeLog_("system","ERROR","Server",where,String(error&&error.message?error.message:error),"",""); }catch(ignore){} }
function friendlyError_(error){ var m=String(error&&error.message?error.message:"เกิดข้อผิดพลาด"); if(m.length>300)m=m.substring(0,300); return m; }
function ok_(data){ return {success:true,data:data}; }
function fail_(message){ return {success:false,message:String(message||"เกิดข้อผิดพลาด")}; }
function nowIso_(){ return new Date().toISOString(); }
function timestampForFile_(){ return Utilities.formatDate(new Date(),Session.getScriptTimeZone()||"Asia/Bangkok","yyyy-MM-dd_HHmm"); }
function round2_(n){ return Math.round(Number(n)*100)/100; }
function numberOrNull_(v){ if(v===""||v===null||v===undefined)return null; var n=Number(v); return isFinite(n)?n:null; }
function cleanText_(v,maxLen){ var s=String(v===undefined||v===null?"":v).replace(/[\u0000-\u001F]/g," ").trim(); if(maxLen&&s.length>maxLen)s=s.substring(0,maxLen); return s; }
function clone_(o){ return safeJson_(JSON.stringify(o),{}); }
function findOne_(arr,key,value){ arr=safeArray_(arr); var i; for(i=0;i<arr.length;i++)if(String(arr[i][key])===String(value))return arr[i]; return null; }
function filterBy_(arr,fn){ arr=safeArray_(arr); var out=[]; var i; for(i=0;i<arr.length;i++)if(fn(arr[i],i))out.push(arr[i]); return out; }
function mapBy_(arr,fn){ arr=safeArray_(arr); var out=[]; var i; for(i=0;i<arr.length;i++)out.push(fn(arr[i],i)); return out; }
function countWhere_(arr,fn){ return filterBy_(arr,fn).length; }
function countBy_(arr,key){ arr=safeArray_(arr); var out={}; var i; for(i=0;i<arr.length;i++){var k=String(arr[i][key]||""); if(!k)continue; out[k]=(out[k]||0)+1;} return out; }
function topKey_(obj){ var best="",max=-1,k; for(k in obj)if(obj.hasOwnProperty(k)&&obj[k]>max){max=obj[k];best=k;} return best; }
function sumField_(arr,key){ arr=safeArray_(arr); var s=0,i; for(i=0;i<arr.length;i++){var n=numberOrNull_(arr[i][key]); if(n!==null)s+=n;} return s; }
function averageField_(arr,key){ arr=safeArray_(arr); var s=0,c=0,i; for(i=0;i<arr.length;i++){var n=numberOrNull_(arr[i][key]); if(n!==null){s+=n;c++;}} return c?round2_(s/c):0; }
