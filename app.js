const D = window.HOMIE_DATA;
const $ = (id) => document.getElementById(id);

function fillSelect(el, rows) {
  el.innerHTML = '';
  rows.forEach((r) => el.append(new Option(r.label, r.id)));
}

function makeList(title, rows) {
  return title + '\n' + (rows || []).map((x, i) => `${i + 1}. ${x}`).join('\n');
}

function setupTabs() {
  document.querySelectorAll('.card').forEach((btn) => {
    btn.onclick = () => {
      document.querySelectorAll('.card').forEach((x) => x.classList.remove('active'));
      document.querySelectorAll('.panel').forEach((x) => x.classList.remove('active'));
      btn.classList.add('active');
      $(btn.dataset.panel).classList.add('active');
    };
  });
}

function setupCalculator() {
  $('calcBtn').onclick = () => {
    const income = Number($('income').value || 0);
    const debt = Number($('debt').value || 0);
    const box = $('calcResult');
    if (income <= 0 || debt < 0) {
      box.className = 'result risk';
      box.textContent = 'กรุณากรอกข้อมูลตัวเลขให้ถูกต้อง';
      return;
    }
    const net = income - debt;
    const ratio = debt / income;
    let status = 'ตรวจต่อได้เบื้องต้น';
    box.className = 'result ok';
    if (ratio >= 0.60) { status = 'ภาระหนี้สูง ควรตรวจสอบเพิ่ม'; box.className = 'result risk'; }
    else if (ratio >= 0.40) { status = 'ควรพิจารณาอย่างระมัดระวัง'; box.className = 'result warn'; }
    box.textContent = `รายได้รวม: ${income.toLocaleString()} บาท/เดือน\nภาระหนี้: ${debt.toLocaleString()} บาท/เดือน\nรายได้หลังหักภาระ: ${net.toLocaleString()} บาท/เดือน\nสัดส่วนภาระหนี้: ${(ratio * 100).toFixed(2)}%\n\nสถานะ: ${status}\n\nหมายเหตุ: เป็นการประเมินจาก Prototype`;
  };
}

function setupChecklist() {
  fillSelect($('requestType'), D.requestTypes);
  fillSelect($('borrowerType'), D.borrowerTypes);
  fillSelect($('businessStatus'), D.businessStatuses);
  fillSelect($('occupation'), D.occupations);
  fillSelect($('loanPurpose'), D.loanPurposes);

  $('borrowerType').onchange = () => {
    const on = $('borrowerType').value === 'BT002';
    $('businessStatus').disabled = !on;
    $('occupation').disabled = !on;
  };
  $('borrowerType').onchange();

  $('buildChecklistBtn').onclick = () => {
    const rt = $('requestType').value;
    const bt = $('borrowerType').value;
    const occ = $('occupation').value;
    const purpose = $('loanPurpose').value;
    const blocks = [];

    if (bt === 'BT001') blocks.push(makeList('เอกสารอาชีพประจำ', D.regularIncomeChecklist[rt]));

    if (bt === 'BT002') {
      blocks.push(makeList('เอกสารพื้นฐานอาชีพอิสระ', D.independentCommonChecklist[rt]));
      const detail = D.occSpecificChecklist[occ];
      if (detail) {
        blocks.push(`อาชีพ: ${detail.name}\nลักษณะอาชีพ: ${detail.description}`);
        blocks.push(makeList('เอกสารเฉพาะอาชีพ', detail.documents));
        blocks.push(`แนวทางพิจารณารายได้\n${detail.analysis}`);
        blocks.push(`ข้อควรระวัง\n${detail.warning}`);
      } else {
        const occName = D.occupations.find((x) => x.id === occ)?.label || 'อาชีพที่เลือก';
        blocks.push(`อาชีพ: ${occName}\nเอกสารเฉพาะอาชีพใน Prototype นี้ยังเป็นโครงหลัก`);
        blocks.push(makeList('เอกสารเบื้องต้น', ['บัญชีเงินฝากย้อนหลัง 6 เดือน', 'หลักฐานประกอบกิจการตามลักษณะอาชีพ', 'หลักฐานรายรับหรือภาษี ถ้ามี', 'รูปถ่ายกิจการหรือช่องทางขาย']));
      }
    }

    if (bt === 'BT003') blocks.push(makeList('เอกสารอาชีพอิสระ Premium', ['ใบประกอบวิชาชีพ / หลักฐานวิชาชีพ', 'บัญชีเงินฝากย้อนหลัง 6 เดือน', 'หลักฐานการรับเงินย้อนหลังตามกรณี', 'งบการเงินย้อนหลัง 2 ปี กรณีนิติบุคคล']));

    blocks.push(makeList('เอกสารหลักประกันตามวัตถุประสงค์กู้', D.purposeChecklist[purpose]));
    blocks.push('หมายเหตุ\nChecklist นี้เป็นผลลัพธ์เบื้องต้นจาก Prototype');
    $('checklistResult').className = 'result ok';
    $('checklistResult').textContent = blocks.join('\n\n');
  };

  $('copyChecklistBtn').onclick = () => {
    navigator.clipboard.writeText($('checklistResult').innerText).then(() => alert('คัดลอกแล้ว'));
  };
}

function setupLearning() {
  const cats = ['ทั้งหมด', ...new Set(D.learning.map((x) => x.category))];
  $('learningFilters').innerHTML = cats.map((x, i) => `<button class="chip ${i ? '' : 'active'}" data-cat="${x}">${x}</button>`).join('');
  $('learningFilters').onclick = (e) => {
    if (!e.target.classList.contains('chip')) return;
    document.querySelectorAll('.chip').forEach((x) => x.classList.remove('active'));
    e.target.classList.add('active');
    renderLearning(e.target.dataset.cat);
  };
  renderLearning('ทั้งหมด');
}

function renderLearning(cat) {
  const rows = cat === 'ทั้งหมด' ? D.learning : D.learning.filter((x) => x.category === cat);
  $('learningCards').innerHTML = rows.map((x) => `<article class="learning-card"><span class="tag">${x.category}</span><h3>${x.title}</h3><p>${x.summary}</p></article>`).join('');
}

function setupQA() {
  $('qaBtn').onclick = () => {
    const q = $('qaInput').value.toLowerCase();
    const found = D.qa.find((item) => item.keys.some((k) => q.includes(k.toLowerCase())));
    $('qaResult').className = found ? 'result ok' : 'result warn';
    $('qaResult').textContent = found ? `คำตอบ\n${found.answer}\n\nหมายเหตุ: ใช้ประกอบการปฏิบัติงานเบื้องต้น` : 'ยังไม่พบคำตอบที่ชัดเจนในฐานข้อมูล Prototype';
  };
}

setupTabs();
setupCalculator();
setupChecklist();
setupLearning();
setupQA();
