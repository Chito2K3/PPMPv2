/**
 * PPMP V2 — Live Network Concurrency Test Runner
 * 
 * Simultaneously dispatches 3 evaluations via HTTP POST to the live Google Apps Script Web App:
 * 1. End-user (Menard Arellano: menardarellano@gmail.com)
 * 2. Pharmacist (Chito Saba: chitosaba@gmail.com)
 * 3. Nurse (Myracel Joy L. Mendoza: myracel.joy@gmail.com)
 * 
 * Target: Identical Medicine & Supplier
 */

const WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzf5cK5DCLZnBB2sE0Cujfwlbbcgc8O84to5XQ--0uVUyrX8zz2D2cD2wIOE4oS1VtG/exec";

const TARGET_MEDICINE = "0.9% Sodium Chloride 1L Bottle/Bag (PNSS 1L)";
const TARGET_SUPPLIER = "2EZ TRADING OPC";
const BRAND_NAME = "SoluPack PNSS";
const MANUFACTURER = "Baxter Healthcare / Local Pharma Corp";
const PRICE = 85.50;

console.log('='.repeat(70));
console.log('🚀 PPMP V2 LIVE NETWORK CONCURRENCY DISPATCHER');
console.log('='.repeat(70));
console.log(`📡 Target Endpoint: ${WEBAPP_URL}`);
console.log(`💊 Medicine:        ${TARGET_MEDICINE}`);
console.log(`🏢 Supplier:        ${TARGET_SUPPLIER}`);
console.log(`💰 Price:           ₱${PRICE.toFixed(2)}\n`);

const dummySig = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

function buildPayload(role, name, email, p1Override, p2Override) {
  const payload = {
    Generic_Name: TARGET_MEDICINE,
    Brand_Name: BRAND_NAME,
    Supplier: TARGET_SUPPLIER,
    Manufacturer: MANUFACTURER,
    Price: PRICE,
    Requires_Reconstitution: "No",
    Recommendation: "Recommended",
    Remarks: `[CONCURRENCY_TEST_SUITE] Live concurrent evaluation by ${name} (${role})`,
    Data_Privacy_Consent: "Yes",
    Accuracy_Consent: "Yes",
    Evaluator_Signature: dummySig,
    evaluatorName: name,
    evaluatorEmail: email,
    Evaluator: role
  };

  for (let i = 1; i <= 19; i++) {
    const num = i < 10 ? `0${i}` : `${i}`;
    payload[`P1_${num}`] = (p1Override && p1Override[`P1_${num}`]) ? p1Override[`P1_${num}`] : "Yes";
  }

  for (let j = 1; j <= 6; j++) {
    const num2 = `0${j}`;
    payload[`P2_${num2}`] = (p2Override && p2Override[`P2_${num2}`]) ? p2Override[`P2_${num2}`] : "Yes";
  }

  return payload;
}

const p1 = buildPayload("End-user", "Menard Arellano", "menardarellano@gmail.com");
const p2 = buildPayload("Pharmacist", "Chito Saba", "chitosaba@gmail.com", { "P1_08": "N/A" });
const p3 = buildPayload("Nurse", "Myracel Joy L. Mendoza", "myracel.joy@gmail.com", null, { "P2_05": "No" });

async function sendEvaluation(payload, label) {
  const start = performance.now();
  try {
    const response = await fetch(WEBAPP_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
      redirect: "follow"
    });

    const elapsed = (performance.now() - start).toFixed(0);
    const text = await response.text();
    let json = null;
    try {
      json = JSON.parse(text);
    } catch (e) {
      // Return raw html/text error if not JSON
      return { label, success: false, status: response.status, raw: text, elapsed };
    }

    return { label, success: json.success, data: json, elapsed };
  } catch (err) {
    return { label, success: false, error: err.message };
  }
}

async function runLiveConcurrencyTest() {
  console.log('⚡ Dispatched 3 evaluations simultaneously via Promise.all()...');
  const t0 = performance.now();

  const [res1, res2, res3] = await Promise.all([
    sendEvaluation(p1, "Agent 1 (End-user)"),
    sendEvaluation(p2, "Agent 2 (Pharmacist)"),
    sendEvaluation(p3, "Agent 3 (Nurse)")
  ]);

  const totalMs = (performance.now() - t0).toFixed(0);
  console.log(`⏱️ All requests completed in ${totalMs} ms.\n`);

  console.log('='.repeat(70));
  console.log('📋 LIVE RESPONSE SUMMARY:');
  console.log('='.repeat(70));

  [res1, res2, res3].forEach(r => {
    console.log(`[${r.label}] Latency: ${r.elapsed} ms`);
    if (r.success) {
      console.log(`  ✅ Success:       ${r.success}`);
      console.log(`  🆔 Evaluation ID: ${r.data.evaluationId}`);
      console.log(`  💬 Message:       ${r.data.message || 'Saved'}`);
    } else {
      console.log(`  ❌ Failed:        ${r.error || (r.raw && r.raw.substring(0, 150)) || 'Unknown Error'}`);
    }
    console.log('');
  });

  if (res1.success && res2.success && res3.success) {
    const ids = [res1.data.evaluationId, res2.data.evaluationId, res3.data.evaluationId];
    const unique = new Set(ids);
    console.log('='.repeat(70));
    console.log('🔍 VERIFICATION ASSERTIONS:');
    console.log('='.repeat(70));
    console.log(`[A1] ✅ Submissions Ingested: 3/3 accepted by Google Sheets backend`);
    console.log(`[A2] ${unique.size === 3 ? '✅' : '❌'} UUID Collision Check: ${unique.size === 3 ? '3 unique UUIDs generated' : 'Duplicate UUID collision!'}`);
    console.log(`     - End-user:   ${ids[0]}`);
    console.log(`     - Pharmacist: ${ids[1]}`);
    console.log(`     - Nurse:      ${ids[2]}`);
    console.log('\n🎉 ALL 3 EVALUATIONS ARE NOW RECORDED IN YOUR GOOGLE SHEET!');
    console.log('Check your Evaluations_Master, Consolidated_Summary, and Evaluator tabs.');
  } else {
    console.log('⚠️ If the response says "Hindi mahanap ang script function: doPost", please ensure');
    console.log('you have updated Code.gs in Apps Script and deployed a New Version.');
  }
  console.log('='.repeat(70));
}

runLiveConcurrencyTest();
