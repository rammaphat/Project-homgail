/**
 * ============================================================================
 * COFFEE MACHINE RECOMMENDATION ENGINE (100-Point MCDM Model)
 * Purely driven by external Excel dataset (products.xlsx) via SheetJS
 * ============================================================================
 */

// 1. DYNAMIC CATALOG STORAGE (NO HARDCODED 20-ITEM ARRAY)
let catalog = [];

// 2. PRESETS CONFIGURATION
const PRESETS = {
  homeBeginner: {
    name: "Home Beginner",
    budget: 15000,
    requiredCups: 8,
    segment: "Home",
    powerSupport: "Standard",
    requireDualBoiler: false,
    requireDirectWater: false,
    description: "Budget-conscious home barista exploring quality espresso drinks."
  },
  prosumer: {
    name: "Prosumer Enthusiast",
    budget: 45000,
    requiredCups: 35,
    segment: "Home",
    powerSupport: "Standard",
    requireDualBoiler: false,
    requireDirectWater: false,
    description: "Default balanced prosumer setup for consistent daily artisan coffee."
  },
  office: {
    name: "Boutique Office",
    budget: 95000,
    requiredCups: 60,
    segment: "Office",
    powerSupport: "Standard",
    requireDualBoiler: false,
    requireDirectWater: true,
    description: "Busy office team needing plumbed water line and high reliability."
  },
  cafeSpecialty: {
    name: "Specialty Cafe",
    budget: 200000,
    requiredCups: 150,
    segment: "Cafe",
    powerSupport: "Standard",
    requireDualBoiler: true,
    requireDirectWater: true,
    description: "Specialty cafe serving 100-180 cups/day with commercial precision."
  },
  commercialHighVolume: {
    name: "Commercial Flagship",
    budget: 400000,
    requiredCups: 350,
    segment: "Cafe",
    powerSupport: "Heavy",
    requireDualBoiler: true,
    requireDirectWater: true,
    description: "High-volume commercial multi-group cafe setup with 3-phase heavy power."
  }
};

// 3. BILINGUAL DICTIONARY
const I18N = {
  th: {
    header_badge: "ระบบวิเคราะห์การตัดสินใจ MCDM",
    header_title: "ระบบค้นหาและแนะนำเครื่องชงกาแฟอัจฉริยะ",
    btn_copy_report: "คัดลอกสรุปผล",
    btn_export_json: "ส่งออก JSON",
    presets_label: "พรีเซ็ตสำเร็จรูป:",
    presets_sub: "คลิกเพื่อโหลดเงื่อนไขตามกลุ่มผู้ใช้งานทันที",
    preset_home: "🌱 บาริสต้ามือใหม่ (฿15k)",
    preset_prosumer: "☕ มาตรฐาน Prosumer (฿45k)",
    preset_office: "🏢 ออฟฟิศขนาดกลาง (฿95k)",
    preset_cafe: "✨ คาเฟ่ระดับ Specialty (฿200k)",
    preset_comm: "⚡ คาเฟ่ปริมาณสูง Commercial (฿400k)",
    hero_badge: "โมเดลตัดสินใจแบบหลายเกณฑ์ (MCDM 100 คะแนน)",
    hero_title: "ระบบคัดเลือกเครื่องชงกาแฟที่ตรงใจและแม่นยำที่สุด",
    hero_desc: "ประเมินและวิเคราะห์เครื่องชงกาแฟจากไฟล์ products.xlsx ด้วยโมเดลคะแนน 100 คะแนน ผสานระบบคัดกรองเงื่อนไขขั้นต้น (Hard Constraints) และ 5 มิติน้ำหนักประเมินคุณค่าอย่างแม่นยำ",
    dim_budget: "ความเหมาะสมงบประมาณ",
    dim_capacity: "จุดคุ้มค่ากำลังผลิต",
    dim_usage: "ตรงกลุ่มการใช้งาน",
    dim_power: "ระบบไฟฟ้าที่รองรับ",
    dim_features: "ฟังก์ชันหม้อต้มคู่/ต่อน้ำตรง",
    param_badge: "การตั้งค่าความต้องการ",
    param_title: "เงื่อนไขของคุณ",
    btn_reset: "คืนค่าเริ่มต้น",
    label_budget: "1. งบประมาณสูงสุด (฿)",
    label_budget_note: "*เกณฑ์เด็ดขาด (Hard Constraint): สินค้าที่ราคาสูงเกิน +15% ของงบจะถูกตัดออกทันที",
    label_cups: "2. จำนวนแก้วที่ต้องชงต่อวัน",
    label_cups_note: "*เกณฑ์เด็ดขาด: เครื่องที่รองรับแก้วต่อวันน้อยกว่าที่ต้องการจะถูกตัดออกทันที (จุดคุ้มค่าสูงสุดคือ 50-85% ของกำลังผลิต)",
    label_segment: "3. สภาพแวดล้อมและรูปแบบการใช้งาน",
    seg_home: "บ้าน",
    seg_home_sub: "ส่วนตัว / โฮมบาริสต้า",
    seg_office: "ออฟฟิศ",
    seg_office_sub: "สตูดิโอ / ทีมงาน",
    seg_cafe: "คาเฟ่",
    seg_cafe_sub: "ร้านกาแฟ / เชิงพาณิชย์",
    label_power: "4. ระบบไฟฟ้าที่รองรับ",
    power_std: "ไฟฟ้ามาตรฐาน",
    power_std_sub: "220V 1 เฟส (ปลั๊กไฟบ้านทั่วไป 10A–16A)",
    power_hvy: "ไฟฟ้าแรงสูงเชิงพาณิชย์",
    power_hvy_sub: "3 เฟส / 20A+ (ไฟสำหรับร้านกาแฟพาณิชย์)",
    power_note: "*เกณฑ์เด็ดขาด: เครื่องระดับ Commercial หนักจะถูกตัดออกหากเลือกระบบไฟฟ้ามาตรฐาน",
    label_features: "5. ฟังก์ชันพิเศษที่ต้องการ (+15 คะแนน)",
    check_db: "ต้องการหม้อต้มคู่ Dual Boiler (+7.5 คะแนน)",
    check_db_sub: "สกัดกาแฟและสตรีมนมพร้อมกันได้ต่อเนื่อง ควบคุมอุณหภูมินิ่งแม่นยำ",
    check_dw: "ต้องการต่อท่อน้ำตรง Direct Water Line (+7.5 คะแนน)",
    check_dw_sub: "ปั๊มโรตารี่ต่อท่อน้ำดีโดยตรง ไม่ต้องคอยเติมน้ำในแทงก์",
    metric_top_score: "ความเหมาะสมสูงสุด",
    metric_eligible: "รุ่นที่ผ่านเกณฑ์",
    metric_disqualified: "รุ่นที่ไม่ผ่านเกณฑ์",
    metric_avg_price: "ราคาเฉลี่ยรุ่นที่ผ่าน",
    top3_badge: "อันดับความเหมาะสมสูงสุด",
    top3_title: "Top 3 เครื่องชงกาแฟที่แนะนำที่สุด",
    top3_sub: "ประเมินและคัดเลือกแบบเรียลไทม์ตามเกณฑ์คะแนน",
    catalog_table_title: "ตารางสรุปผลการจัดอันดับทั้งหมด (Full Catalog Ranking)",
    catalog_table_sub: "เรียงตามคะแนนความเหมาะสม MCDM สูงสุดลงมา",
    sort_label: "เรียงตาม:",
    sort_score: "คะแนนสูงสุด (แนะนำ)",
    sort_price_asc: "ราคา: ต่ำไปสูง",
    sort_price_desc: "ราคา: สูงไปต่ำ",
    sort_cups_desc: "กำลังผลิต: มากไปน้อย",
    filter_all_seg: "ทุกกลุ่มการใช้งาน",
    filter_home: "เฉพาะบ้าน (Home)",
    filter_office: "เฉพาะออฟฟิศ (Office)",
    filter_cafe: "เฉพาะร้านกาแฟ (Cafe)",
    filter_all_brands: "ทุกแบรนด์",
    th_rank: "อันดับ",
    th_product: "ชื่อรุ่นและยี่ห้อ",
    th_segment: "กลุ่มใช้งาน",
    th_power: "ระบบไฟ",
    th_features: "ฟังก์ชัน",
    th_price: "ราคาจำหน่าย",
    th_capacity: "กำลังผลิต/ภาระงาน",
    th_score: "ความเหมาะสม (Match Rate)",
    th_actions: "การจัดการ",
    disq_title: "รุ่นที่ไม่ผ่านเกณฑ์และข้อจำกัดที่ไม่รองรับ",
    disq_view_reasons: "ดูเหตุผลที่ไม่ผ่าน",
    disq_desc: "รุ่นต่อไปนี้ไม่ผ่านเกณฑ์บังคับขั้นต้น (เกินงบ 115%, ไฟฟ้าไม่รองรับ, หรือความจุไม่พอ) ตามหลัก MCDM จึงได้รับ 0 คะแนน",
    compare_dock_title: "เปรียบเทียบเครื่อง:",
    btn_view_compare: "เปิดตารางเปรียบเทียบ",
    btn_specs_breakdown: "ดูสเปกและคะแนน",
    btn_add_compare: "เปรียบเทียบ",
    btn_added: "เลือกแล้ว",
    points_unit: "คะแนน",
    budget_col: "งบประมาณ",
    capacity_col: "ความจุ",
    usage_col: "การใช้งาน",
    power_col: "ระบบไฟ",
    features_col: "ฟังก์ชัน",
    disqualified_badge: "ไม่ผ่านเกณฑ์ (0 คะแนน)",
    no_results_found: "ไม่พบเครื่องชงกาแฟที่ตรงตามเงื่อนไข",
    no_results_desc: "เครื่องทั้งหมดไม่ผ่านเงื่อนไขที่คุณเลือก ลองเพิ่มงบประมาณ ลดจำนวนแก้วที่ต้องการ หรือเลือกระบบไฟฟ้าแรงสูง"
  },
  en: {
    header_badge: "MCDM Decision Engine",
    header_title: "Coffee Machine Recommender",
    btn_copy_report: "Copy Report",
    btn_export_json: "Export JSON",
    presets_label: "Quick Presets:",
    presets_sub: "Instant persona setups with calibrated criteria",
    preset_home: "🌱 Home Beginner (฿15k)",
    preset_prosumer: "☕ Prosumer Standard (฿45k)",
    preset_office: "🏢 Boutique Office (฿95k)",
    preset_cafe: "✨ Specialty Cafe (฿200k)",
    preset_comm: "⚡ High-Volume Commercial (฿400k)",
    hero_badge: "Multi-Criteria Decision Making (MCDM)",
    hero_title: "Precision Coffee Machine Matchmaker",
    hero_desc: "Evaluating coffee machines loaded directly from products.xlsx using a 100-point multi-attribute model with strict pre-filter constraints and 5 weighted dimensions.",
    dim_budget: "Budget Match",
    dim_capacity: "Capacity Sweetspot",
    dim_usage: "Usage Segment",
    dim_power: "Power Match",
    dim_features: "Features (DB & Water)",
    param_badge: "Parameters",
    param_title: "Your Requirements",
    btn_reset: "Reset",
    label_budget: "1. Maximum Budget (฿)",
    label_budget_note: "*Hard Constraint: Products over +15% of this budget are strictly eliminated.",
    label_cups: "2. Required Daily Cups",
    label_cups_note: "*Hard Constraint: Machines with lower max rated cups are disqualified. Sweet spot is 50-85% machine load.",
    label_segment: "3. Intended Environment & Usage Segment",
    seg_home: "Home",
    seg_home_sub: "Personal / Barista",
    seg_office: "Office",
    seg_office_sub: "Studio / Team",
    seg_cafe: "Cafe",
    seg_cafe_sub: "Specialty / Commercial",
    label_power: "4. Electrical Infrastructure",
    power_std: "Standard Power",
    power_std_sub: "220V 1-phase (10A–16A standard outlet)",
    power_hvy: "Heavy Duty Power",
    power_hvy_sub: "3-Phase / 20A+ commercial electrical grid",
    power_note: "*Hard Constraint: Commercial 'Heavy' machines are eliminated if Standard power is selected.",
    label_features: "5. Specific Feature Match (+15 pts)",
    check_db: "Require Dual Boiler (+7.5 pts)",
    check_db_sub: "Simultaneous espresso brewing & continuous high-pressure milk steaming without temperature wait.",
    check_dw: "Require Direct Water Line Plumbing (+7.5 pts)",
    check_dw_sub: "Rotary pump connected directly to mains water filtration supply (no manual tank refilling).",
    metric_top_score: "Top Match Rate",
    metric_eligible: "Qualified Models",
    metric_disqualified: "Disqualified Models",
    metric_avg_price: "Average Qualified Price",
    top3_badge: "Top Match Recommendations",
    top3_title: "Top 3 Recommended Machines",
    top3_sub: "Evaluated dynamically by MCDM model",
    catalog_table_title: "Full Catalog Ranking",
    catalog_table_sub: "All passing products sorted by Total MCDM Score descending",
    sort_label: "Sort By:",
    sort_score: "Highest Score (Default)",
    sort_price_asc: "Price: Low to High",
    sort_price_desc: "Price: High to Low",
    sort_cups_desc: "Capacity: High to Low",
    filter_all_seg: "All Segments",
    filter_home: "Home Only",
    filter_office: "Office Only",
    filter_cafe: "Cafe Only",
    filter_all_brands: "All Brands",
    th_rank: "Rank",
    th_product: "Product & Brand",
    th_segment: "Segment",
    th_power: "Power",
    th_features: "Specs",
    th_price: "Price",
    th_capacity: "Capacity",
    th_score: "Match Rate",
    th_actions: "Actions",
    disq_title: "Disqualified & Incompatible Models",
    disq_view_reasons: "View Reasons",
    disq_desc: "The following models failed your strict pre-filter constraints (exceeded 115% budget, required heavy 3-phase electricity, or insufficient daily cup capacity). In accordance with the MCDM rules, they receive 0 total points.",
    compare_dock_title: "Compare Models:",
    btn_view_compare: "View Comparison Matrix",
    btn_specs_breakdown: "Specs & Breakdown",
    btn_add_compare: "Compare",
    btn_added: "Added",
    points_unit: "Points",
    budget_col: "Budg",
    capacity_col: "Cap",
    usage_col: "Use",
    power_col: "Pwr",
    features_col: "Feat",
    disqualified_badge: "DISQUALIFIED (0 pts)",
    no_results_found: "No Matching Machines Found",
    no_results_desc: "All models were filtered out by your current constraints. Try increasing your budget, lowering required cups, or enabling Heavy electrical support."
  }
};

// 4. APPLICATION STATE (Defaults: 45,000 THB Budget, 35 Daily Cups)
const state = {
  lang: "th",
  user: {
    maxBudget: 45000,
    requiredCups: 35,
    segment: "Home",
    powerSupport: "Standard",
    requireDualBoiler: false,
    requireDirectWater: false
  },
  filters: {
    sortBy: "score_desc",
    segment: "all",
    brand: "all",
    showDisqualified: true
  },
  activePreset: "prosumer",
  comparedModelIds: [],
  evaluatedCatalog: []
};

// 5. EXCEL DATA LOADING & NORMALIZATION (USING SHEETJS)

/**
 * Normalizes varied spreadsheet column names (English, Thai, with/without whitespace)
 */
function normalizeRow(row, index) {
  const norm = {
    id: `CM-${String(index + 1).padStart(3, '0')}`,
    name: "",
    brand: "Generic",
    model: "",
    price: 0,
    maxCups: 10,
    segment: "Home",
    power: "Standard",
    dualBoiler: false,
    directWater: false
  };

  for (const rawKey of Object.keys(row)) {
    const key = rawKey.trim().toLowerCase();
    const val = row[rawKey];

    // ID
    if (key === 'id' || key.includes('product id') || key.includes('รหัส')) {
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        norm.id = String(val).trim();
      }
    }
    // Brand & Model / Product Name
    else if (key.includes('brand') || key.includes('model') || key.includes('รุ่น') || key.includes('สินค้า') || key.includes('ชื่อ')) {
      norm.name = String(val).trim();
    }
    // Price
    else if (key.includes('price') || key.includes('ราคา')) {
      norm.price = Number(String(val).replace(/[^0-9.-]+/g, '')) || 0;
    }
    // MaxCups
    else if (key.includes('cups') || key.includes('แก้ว') || key.includes('สูงสุด') || key.includes('max')) {
      norm.maxCups = Number(String(val).replace(/[^0-9.-]+/g, '')) || 0;
    }
    // Segment (Home, Office, Cafe)
    else if (key.includes('segment') || key.includes('ประเภท') || key.includes('กลุ่ม')) {
      const segStr = String(val).trim();
      if (/office/i.test(segStr) || /ออฟฟิศ/i.test(segStr) || /สำนักงาน/i.test(segStr)) {
        norm.segment = 'Office';
      } else if (/cafe/i.test(segStr) || /คาเฟ่/i.test(segStr) || /ร้านกาแฟ/i.test(segStr)) {
        norm.segment = 'Cafe';
      } else {
        norm.segment = 'Home';
      }
    }
    // Power (Standard, Heavy)
    else if (key.includes('power') || key.includes('ไฟ') || key.includes('ระบบไฟ')) {
      const pwrStr = String(val).trim();
      norm.power = (/heavy/i.test(pwrStr) || /3/i.test(pwrStr) || /แรงสูง/i.test(pwrStr)) ? 'Heavy' : 'Standard';
    }
    // Dual Boiler (Boolean)
    else if (key.includes('dual') || key.includes('boiler') || key.includes('หม้อต้มคู่')) {
      norm.dualBoiler = (val === 1 || val === true || String(val).trim().toLowerCase() === 'true' || String(val).trim() === '1');
    }
    // Direct Water Line (Boolean)
    else if (key.includes('direct') || key.includes('water') || key.includes('น้ำ') || key.includes('ตรง') || key.includes('plumb')) {
      norm.directWater = (val === 1 || val === true || String(val).trim().toLowerCase() === 'true' || String(val).trim() === '1');
    }
  }

  // Brand & Model Extraction
  const knownBrands = [
    'La Marzocco', 'Nuova Simonelli', "De'Longhi", 'Breville', 
    'Rancilio', 'Lelit', 'Profitec', 'Rocket', 'Sanremo', 'Gaggia', 'Flair'
  ];
  let brand = 'Generic';
  let model = norm.name || norm.id;

  for (const b of knownBrands) {
    if (norm.name && norm.name.toLowerCase().startsWith(b.toLowerCase())) {
      brand = b;
      model = norm.name.substring(b.length).trim() || norm.name;
      break;
    }
  }

  norm.brand = brand;
  norm.model = model;
  return norm;
}

/**
 * Loads and parses products.xlsx via fetch() using SheetJS (XLSX)
 */
async function loadCatalogFromExcel() {
  const statusEl = document.getElementById("excel-status-text");
  const errorBanner = document.getElementById("excel-error-banner");

  if (statusEl) {
    statusEl.textContent = state.lang === 'th' ? "กำลังโหลด products.xlsx..." : "Loading products.xlsx...";
  }

  try {
    let arrayBuffer;

    // Node.js environment support for automated testing
    if (typeof window === 'undefined' && typeof require !== 'undefined') {
      const fs = require('fs');
      const filename = fs.existsSync('products.xlsx') ? 'products.xlsx' : 'product.xlsx';
      arrayBuffer = fs.readFileSync(filename);
    } else {
      // Browser environment: fetch from root directory
      let response = await fetch('./products.xlsx');
      if (!response.ok) {
        // Fallback check for product.xlsx
        response = await fetch('./product.xlsx');
      }
      if (!response.ok) {
        throw new Error(`Failed to fetch Excel file (HTTP ${response.status})`);
      }
      arrayBuffer = await response.arrayBuffer();
    }

    // Parse workbook via SheetJS
    const XLSXLib = (typeof XLSX !== 'undefined') ? XLSX : (typeof require !== 'undefined' ? require('xlsx') : null);
    if (!XLSXLib) {
      throw new Error("SheetJS (XLSX) library is not loaded.");
    }

    const workbook = XLSXLib.read(arrayBuffer, { type: (typeof window === 'undefined' ? 'buffer' : 'array') });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawJson = XLSXLib.utils.sheet_to_json(worksheet);

    if (!rawJson || rawJson.length === 0) {
      throw new Error("Spreadsheet contains no data rows.");
    }

    // Normalize catalog records
    catalog = rawJson.map((row, idx) => normalizeRow(row, idx));

    // Update status badge
    if (statusEl) {
      statusEl.textContent = state.lang === 'th'
        ? `${catalog.length} รุ่นจาก products.xlsx`
        : `${catalog.length} machines from products.xlsx`;
    }
    if (errorBanner) {
      errorBanner.classList.add("hidden");
    }

    // Refresh UI
    renderUI();
  } catch (error) {
    console.warn("Excel loading note:", error);
    if (statusEl) {
      statusEl.textContent = state.lang === 'th' ? "ต้องเลือกไฟล์ products.xlsx" : "Awaiting products.xlsx";
    }
    if (errorBanner) {
      errorBanner.classList.remove("hidden");
      const errorMsg = document.getElementById("excel-error-msg");
      if (errorMsg) {
        errorMsg.textContent = state.lang === 'th'
          ? "ไม่สามารถ auto-fetch ไฟล์ Excel ได้ (พบบ่อยเมื่อเปิดไฟล์ HTML ตรงๆ โดยไม่ผ่าน HTTP Server) โปรดคลิก 'Select products.xlsx' ด้านล่างเพื่อโหลดไฟล์ครับ"
          : `Auto-fetch failed (${error.message}). If opening via local file://, please select products.xlsx using the button below.`;
      }
    }
  }
}

/**
 * Parses user-selected file from local file input (fallback for file:// origin)
 */
function handleManualFileSelection(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const XLSXLib = (typeof XLSX !== 'undefined') ? XLSX : null;
      if (!XLSXLib) throw new Error("SheetJS not available");
      const workbook = XLSXLib.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawJson = XLSXLib.utils.sheet_to_json(worksheet);

      catalog = rawJson.map((row, idx) => normalizeRow(row, idx));

      const statusEl = document.getElementById("excel-status-text");
      const errorBanner = document.getElementById("excel-error-banner");
      if (statusEl) {
        statusEl.textContent = state.lang === 'th'
          ? `${catalog.length} รุ่นจาก ${file.name}`
          : `${catalog.length} machines from ${file.name}`;
      }
      if (errorBanner) errorBanner.classList.add("hidden");

      renderUI();
    } catch (err) {
      alert("Error reading file: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// 6. CORE SCORING & MCDM ENGINE

/**
 * Hard Constraints (Pre-filter)
 * Filter out products from scoring if:
 * 1. user.requiredCups > product.maxCups
 * 2. user.powerSupport === 'Standard' AND product.power === 'Heavy'
 * 3. product.price > user.maxBudget * 1.15
 */
function evaluateConstraints(product, user) {
  const violations = [];
  const isTh = typeof state !== 'undefined' && state.lang === 'th';

  if (user.requiredCups > product.maxCups) {
    violations.push({
      code: "CAPACITY_DEFICIT",
      title: isTh ? "กำลังผลิตไม่เพียงพอ" : "Capacity Deficit",
      description: isTh
        ? `รองรับได้สูงสุด ${product.maxCups} แก้ว/วัน แต่ต้องการ ${user.requiredCups} แก้ว/วัน`
        : `Rated for max ${product.maxCups} cups/day, but you require ${user.requiredCups} cups/day.`
    });
  }

  if (user.powerSupport === 'Standard' && product.power === 'Heavy') {
    violations.push({
      code: "POWER_MISMATCH",
      title: isTh ? "ระบบไฟฟ้าไม่รองรับ" : "Power Mismatch",
      description: isTh
        ? "เครื่องต้องใช้ไฟแรงสูง Heavy (3 เฟส/แอมป์สูง) แต่เลือกใช้ระบบ Standard 15A"
        : "Requires Heavy power (3-phase/high amp), but user selected Standard 15A."
    });
  }

  const budgetCeiling = user.maxBudget * 1.15;
  if (product.price > budgetCeiling) {
    const overPct = Math.round(((product.price - user.maxBudget) / user.maxBudget) * 100);
    violations.push({
      code: "BUDGET_EXCEEDED",
      title: isTh ? "เกินงบประมาณที่กำหนด" : "Exceeds Budget Limit",
      description: isTh
        ? `ราคา ฿${product.price.toLocaleString()} สูงเกินงบไป ${overPct}% (เกณฑ์เพดานคือ +15%)`
        : `Price ฿${product.price.toLocaleString()} is ${overPct}% over budget (limit is +15%).`
    });
  }

  return {
    isEligible: violations.length === 0,
    violations
  };
}

/**
 * Scoring Architecture (100 Points Total)
 * Evaluates passing products across 5 weighted dimensions:
 * 1. Budget Match (30 pts)
 * 2. Capacity Match (25 pts)
 * 3. Usage Match (20 pts)
 * 4. Power Match (10 pts)
 * 5. Feature Match (15 pts)
 */
function calculateMCDM(product, user, isEligible) {
  const isTh = typeof state !== 'undefined' && state.lang === 'th';

  if (!isEligible) {
    return {
      budgetScore: 0,
      capacityScore: 0,
      usageScore: 0,
      powerScore: 0,
      featureScore: 0,
      dualBoilerSubScore: 0,
      directWaterSubScore: 0,
      totalScore: 0,
      ratio: user.requiredCups / (product.maxCups || 1),
      priceRatio: product.price / (user.maxBudget || 1),
      verdict: isTh ? "ไม่ผ่านเกณฑ์ขั้นต้น (Hard Constraint)" : "Disqualified by hard constraint"
    };
  }

  // 1. Budget Match (30 pts)
  let budgetScore = 0;
  const priceRatio = product.price / user.maxBudget;
  let budgetNote = "";

  if (product.price <= user.maxBudget) {
    if (priceRatio >= 0.8) {
      budgetScore = 30; // 80% to 100% of budget -> 30 pts
      budgetNote = isTh ? "จัดสรรงบประมาณได้คุ้มค่าสูงสุด (80-100%)" : "Optimal budget allocation (80-100%)";
    } else if (priceRatio >= 0.5) {
      budgetScore = 25; // 50% to 79% of budget -> 25 pts
      budgetNote = isTh ? "ใช้งบกำลังดี มีเงินสำรองเหลือ (50-79%)" : "Comfortably under budget (50-79%)";
    } else {
      budgetScore = 20; // < 50% of budget -> 20 pts
      budgetNote = isTh ? "ต่ำกว่างบมาก ประหยัดต้นทุน (< 50%)" : "Substantially under budget (< 50%)";
    }
  } else if (product.price <= user.maxBudget * 1.15) {
    budgetScore = 10; // Up to 1.15x budget -> 10 pts
    budgetNote = isTh ? "เกินงบเล็กน้อย (อยู่ในเกณฑ์ยืดหยุ่นไม่เกิน +15%)" : "Within +15% stretch tolerance";
  } else {
    budgetScore = 0;
    budgetNote = isTh ? "เกินงบประมาณที่กำหนด" : "Exceeds budget allowance";
  }

  // 2. Capacity Match (25 pts)
  let capacityScore = 0;
  const ratio = user.requiredCups / (product.maxCups || 1);
  let capacityNote = "";

  if (ratio >= 0.5 && ratio <= 0.85) {
    capacityScore = 25; // Sweet spot (50% - 85%)
    capacityNote = isTh ? "จุดคุ้มค่าสูงสุด: ภาระงาน 50-85% ของกำลังผลิต" : "Sweet spot: 50-85% duty cycle";
  } else if (ratio >= 0.3 && ratio < 0.5) {
    capacityScore = 18; // Headroom (30% - 49%)
    capacityNote = isTh ? "มีกำลังผลิตสำรองเหลือเฟือ (30-49% duty cycle)" : "Comfortable headroom (30-49% load)";
  } else if (ratio < 0.3) {
    capacityScore = 12; // Far under (< 30%)
    capacityNote = isTh ? "กำลังผลิตเกินความต้องการมาก (< 30% duty cycle)" : "Excess capacity (< 30% load)";
  } else {
    capacityScore = 0; // Strained (> 85% or > 1.0)
    capacityNote = isTh ? "เครื่องทำงานหนักต่อเนื่อง (> 85% load)" : "High duty strain (> 85% load)";
  }

  // 3. Usage Match (20 pts)
  let usageScore = 0;
  let usageNote = "";

  if (user.segment === product.segment) {
    usageScore = 20; // Exact match
    usageNote = isTh ? `ตรงกับรูปแบบการใช้งาน (${user.segment})` : `Exact match (${user.segment})`;
  } else if (
    (user.segment === "Home" && product.segment === "Office") ||
    (user.segment === "Office" && product.segment === "Cafe")
  ) {
    usageScore = 10; // Adjacent match
    usageNote = isTh ? `อัปเกรดเกรดการใช้งาน (${user.segment} → ${product.segment})` : `Adjacent match (${user.segment} → ${product.segment})`;
  } else {
    usageScore = 0;
    usageNote = isTh ? `ไม่ตรงกลุ่มการใช้งาน (${user.segment} vs ${product.segment})` : `Segment mismatch (${user.segment} vs ${product.segment})`;
  }

  // 4. Power Match (10 pts)
  let powerScore = 0;
  let powerNote = "";

  if (product.power === 'Standard') {
    powerScore = 10;
    powerNote = isTh ? "ระบบไฟ 220V 15A มาตรฐาน เสียบปลั๊กได้ทันที" : "Standard 15A plug-and-play";
  } else if (product.power === 'Heavy' && user.powerSupport === 'Heavy') {
    powerScore = 10;
    powerNote = isTh ? "รองรับไฟฟ้าแรงสูง Heavy 3-Phase" : "Heavy 3-Phase power verified";
  } else {
    powerScore = 0;
    powerNote = isTh ? "ระบบไฟฟ้าไม่ตรงกัน" : "Power mismatch";
  }

  // 5. Machine Fit / Features Match (15 pts)
  let dualBoilerSubScore = 0;
  let dualBoilerNote = "";
  if (user.requireDualBoiler) {
    if (product.dualBoiler) {
      dualBoilerSubScore = 7.5;
      dualBoilerNote = isTh ? "มีหม้อต้มคู่ตามต้องการ (+7.5)" : "Dual boiler fulfilled (+7.5)";
    } else {
      dualBoilerSubScore = 0;
      dualBoilerNote = isTh ? "ไม่มีหม้อต้มคู่ (0/7.5)" : "Missing dual boiler (0/7.5)";
    }
  } else {
    dualBoilerSubScore = 7.5;
    dualBoilerNote = isTh ? "ไม่ต้องการหม้อต้มคู่ (+7.5)" : "Dual boiler not required (+7.5)";
  }

  let directWaterSubScore = 0;
  let directWaterNote = "";
  if (user.requireDirectWater) {
    if (product.directWater) {
      directWaterSubScore = 7.5;
      directWaterNote = isTh ? "มีระบบต่อน้ำตรงตามต้องการ (+7.5)" : "Direct water line fulfilled (+7.5)";
    } else {
      directWaterSubScore = 0;
      directWaterNote = isTh ? "ไม่มีระบบต่อน้ำตรง (0/7.5)" : "Missing direct plumbing (0/7.5)";
    }
  } else {
    directWaterSubScore = 7.5;
    directWaterNote = isTh ? "ไม่ต้องการต่อน้ำตรง (+7.5)" : "Direct plumbing not required (+7.5)";
  }

  const featureScore = dualBoilerSubScore + directWaterSubScore;
  const totalScore = Math.round((budgetScore + capacityScore + usageScore + powerScore + featureScore) * 10) / 10;

  // Short explainability text
  let explainability = "";
  if (isTh) {
    if (capacityScore === 25 && budgetScore >= 25) {
      explainability = `จุดคุ้มค่ากำลังผลิตยอดเยี่ยม (${Math.round(ratio * 100)}% load), จัดสรรงบประมาณได้คุ้มค่า`;
    } else if (capacityScore === 25) {
      explainability = `ความจุกำลังผลิตพอดีกับการใช้งาน (Sweet spot 50-85%), ระบบไฟมาตรฐาน`;
    } else if (budgetScore === 30) {
      explainability = `ใช้ประโยชน์จากงบประมาณได้เต็มประสิทธิภาพ, มีกำลังผลิตรองรับ ${product.maxCups} แก้ว/วัน`;
    } else {
      explainability = `ผ่านเกณฑ์ทุกด้าน, ระบบไฟ ${product.power}, รองรับแก้วต่อวัน ${product.maxCups} แก้ว`;
    }
  } else {
    if (capacityScore === 25 && budgetScore >= 25) {
      explainability = "Optimal volume sweet spot capacity, within ideal budget allocation";
    } else if (capacityScore === 25) {
      explainability = "Optimal volume capacity sweet spot (50-85%), within standard power limits";
    } else if (budgetScore === 30) {
      explainability = `Maximizes allocated budget value, rated for ${product.maxCups} cups/day`;
    } else {
      explainability = `Balanced fit across criteria, ${product.power} power, handles ${product.maxCups} cups/day`;
    }
  }

  return {
    budgetScore,
    budgetNote,
    capacityScore,
    capacityNote,
    usageScore,
    usageNote,
    powerScore,
    powerNote,
    featureScore,
    dualBoilerSubScore,
    dualBoilerNote,
    directWaterSubScore,
    directWaterNote,
    totalScore,
    ratio,
    priceRatio,
    explainability
  };
}

// 7. EVALUATION PIPELINE
function evaluateCatalog() {
  state.evaluatedCatalog = catalog.map(product => {
    const constraintCheck = evaluateConstraints(product, state.user);
    const scoreData = calculateMCDM(product, state.user, constraintCheck.isEligible);

    return {
      ...product,
      isEligible: constraintCheck.isEligible,
      violations: constraintCheck.violations,
      score: scoreData
    };
  });

  sortEvaluatedCatalog();
}

function sortEvaluatedCatalog() {
  const { sortBy } = state.filters;

  state.evaluatedCatalog.sort((a, b) => {
    if (a.isEligible && !b.isEligible) return -1;
    if (!a.isEligible && b.isEligible) return 1;

    switch (sortBy) {
      case "price_asc":
        return a.price - b.price;
      case "price_desc":
        return b.price - a.price;
      case "cups_desc":
        return b.maxCups - a.maxCups;
      case "score_desc":
      default:
        if (b.score.totalScore !== a.score.totalScore) {
          return b.score.totalScore - a.score.totalScore;
        }
        return b.score.budgetScore - a.score.budgetScore;
    }
  });
}

// 8. UI RENDER FUNCTIONS & HUMAN-READABLE MATCH RATES

/**
 * Normalizes raw sub-scores into Human-Readable Match Rates (0–100%),
 * with dynamic color coding and concise Thai micro-copy.
 */
function getHumanScoring(product, user, score) {
  const isTh = (typeof state !== 'undefined' ? state.lang : 'th') === 'th';

  // Overall Match (0–100%)
  const overallPct = Math.round(score.totalScore);
  const overallColor = overallPct >= 80 ? "#22c55e" : (overallPct >= 50 ? "#eab308" : "#ef4444");
  const overallBadgeClass = overallPct >= 80
    ? "match-badge-success"
    : (overallPct >= 50 ? "match-badge-warning" : "match-badge-danger");

  // 1. Budget Match (Max 30) -> 0–100%
  const budgetPct = Math.round((score.budgetScore / 30) * 100);
  let budgetColor = "#22c55e";
  let budgetNote = "";
  if (isTh) {
    if (budgetPct === 100) {
      budgetColor = "#22c55e";
      budgetNote = "คุ้มค่า สเปกเต็มงบประมาณที่ตั้งไว้";
    } else if (budgetPct >= 70) {
      budgetColor = "#22c55e";
      budgetNote = "ประหยัดงบ ได้สเปกพอดีใช้งาน";
    } else if (product.price > user.maxBudget) {
      budgetColor = "#eab308";
      budgetNote = "เกินงบเล็กน้อย แต่สเปกคุ้มราคา";
    } else {
      budgetColor = "#eab308";
      budgetNote = "ประหยัดงบได้มาก";
    }
  } else {
    if (budgetPct === 100) {
      budgetColor = "#22c55e";
      budgetNote = "Optimal value matching full budget";
    } else if (budgetPct >= 70) {
      budgetColor = "#22c55e";
      budgetNote = "Budget-friendly, fits your needs";
    } else if (product.price > user.maxBudget) {
      budgetColor = "#eab308";
      budgetNote = "Slightly over budget, but high value";
    } else {
      budgetColor = "#eab308";
      budgetNote = "Substantial budget savings";
    }
  }

  // 2. Capacity Match (Max 25) -> 0–100%
  const capacityPct = Math.round((score.capacityScore / 25) * 100);
  const capacityColor = capacityPct >= 80 ? "#22c55e" : (capacityPct >= 50 ? "#eab308" : "#ef4444");
  let capacityNote = "";
  if (isTh) {
    if (capacityPct === 100) {
      capacityNote = `รองรับ ${user.requiredCups} แก้ว/วัน ได้ต่อเนื่อง มีกำลังสำรอง`;
    } else if (capacityPct >= 60) {
      capacityNote = `สเปกสูงเกินจำเป็นเล็กน้อย (รองรับได้สูงสุด ${product.maxCups} แก้ว)`;
    } else {
      const loadRatio = user.requiredCups / (product.maxCups || 1);
      if (loadRatio < 0.3) {
        capacityNote = `สเปกสูงเกินจำเป็นมาก (รองรับได้สูงสุด ${product.maxCups} แก้ว)`;
      } else {
        capacityNote = "กำลังผลิตอาจไม่เพียงพอ";
      }
    }
  } else {
    if (capacityPct === 100) capacityNote = `Handles ${user.requiredCups} cups/day smoothly with duty headroom`;
    else if (capacityPct >= 60) capacityNote = `Slightly oversized (max ${product.maxCups} cups/day)`;
    else capacityNote = "Capacity may be strained";
  }

  // 3. Usage Match (Max 20) -> 0–100%
  const usagePct = Math.round((score.usageScore / 20) * 100);
  const usageColor = usagePct >= 80 ? "#22c55e" : (usagePct >= 50 ? "#eab308" : "#ef4444");
  const segMapTh = { Home: "บ้าน", Office: "ออฟฟิศ", Cafe: "คาเฟ่" };
  const targetSegTh = segMapTh[user.segment] || user.segment;
  let usageNote = "";
  if (isTh) {
    if (usagePct === 100) {
      usageNote = `ออกแบบมาเพื่อการใช้งานแบบ${targetSegTh}โดยเฉพาะ`;
    } else if (usagePct >= 50) {
      usageNote = "สามารถประยุกต์ใช้งานได้";
    } else {
      usageNote = "ไม่ตรงกลุ่มการใช้งานหลัก";
    }
  } else {
    if (usagePct === 100) usageNote = `Specifically engineered for ${user.segment} environments`;
    else if (usagePct >= 50) usageNote = "Adaptable for this use-case";
    else usageNote = "Different primary environment";
  }

  // 4. Power Match (Max 10) -> 0–100%
  const powerPct = Math.round((score.powerScore / 10) * 100);
  const powerColor = powerPct >= 80 ? "#22c55e" : (powerPct >= 50 ? "#eab308" : "#ef4444");
  let powerNote = "";
  if (isTh) {
    if (powerPct === 100) {
      powerNote = product.power === "Standard"
        ? "ใช้ไฟบ้านปกติ (15A) เสียบปลั๊กใช้งานได้ทันที"
        : "ระบบไฟกำลังสูง รองรับตามสถานที่ของคุณ";
    } else {
      powerNote = "ระบบไฟฟ้าไม่รองรับ";
    }
  } else {
    if (powerPct === 100) {
      powerNote = product.power === "Standard"
        ? "Standard 15A plug-and-play"
        : "Heavy commercial 3-phase grid compatible";
    } else {
      powerNote = "Incompatible electrical power";
    }
  }

  // 5. Feature Fit (Max 15) -> 0–100%
  const featurePct = Math.round((score.featureScore / 15) * 100);
  const featureColor = featurePct >= 80 ? "#22c55e" : (featurePct >= 50 ? "#eab308" : "#ef4444");
  const featureBadges = [];
  if (product.dualBoiler) {
    featureBadges.push({ label: isTh ? "✓ มี Dual Boiler" : "✓ Dual Boiler", active: true });
  }
  if (product.directWater) {
    featureBadges.push({ label: isTh ? "✓ รองรับต่อน้ำตรง" : "✓ Direct Plumbing", active: true });
  }
  if (featureBadges.length === 0) {
    featureBadges.push({ label: isTh ? "แทงก์น้ำในตัว (ระบบเดี่ยว)" : "Internal Reservoir", active: false });
  }

  return {
    overall: { pct: overallPct, color: overallColor, badgeClass: overallBadgeClass },
    budget: { title: isTh ? "งบประมาณ (Budget)" : "Budget Match", pct: budgetPct, color: budgetColor, note: budgetNote },
    capacity: { title: isTh ? "กำลังผลิต (Capacity)" : "Capacity Match", pct: capacityPct, color: capacityColor, note: capacityNote },
    usage: { title: isTh ? "สภาพแวดล้อม (Usage)" : "Usage Segment", pct: usagePct, color: usageColor, note: usageNote },
    power: { title: isTh ? "ระบบไฟฟ้า (Power)" : "Power Match", pct: powerPct, color: powerColor, note: powerNote },
    features: { title: isTh ? "ฟังก์ชัน (Features)" : "Feature Fit", pct: featurePct, color: featureColor, badges: featureBadges }
  };
}

function getScoreColorClass(score) {
  if (score >= 80) return "text-emerald-700 bg-emerald-50 border-emerald-300";
  if (score >= 50) return "text-amber-700 bg-amber-50 border-amber-300";
  return "text-red-700 bg-red-50 border-red-300";
}

function renderUI() {
  applyTranslations();
  evaluateCatalog();

  const eligibleList = state.evaluatedCatalog.filter(m => m.isEligible);
  const disqualifiedList = state.evaluatedCatalog.filter(m => !m.isEligible);

  // Apply filters
  const filteredEligible = eligibleList.filter(m => {
    if (state.filters.segment !== "all" && m.segment !== state.filters.segment) return false;
    if (state.filters.brand !== "all" && m.brand !== state.filters.brand) return false;
    return true;
  });

  const filteredDisqualified = disqualifiedList.filter(m => {
    if (state.filters.segment !== "all" && m.segment !== state.filters.segment) return false;
    if (state.filters.brand !== "all" && m.brand !== state.filters.brand) return false;
    return true;
  });

  // Render Metrics Summary Bar
  renderMetricsBar(eligibleList, disqualifiedList);

  // Render "Top 3 Recommended Machines" Cards
  // Top 3 must ALWAYS be ranked by Highest Match Rate (Score descending),
  // completely decoupled from the Full Catalog Table's sort filter
  const topRecommendations = [...filteredEligible].sort((a, b) => {
    if (b.score.totalScore !== a.score.totalScore) {
      return b.score.totalScore - a.score.totalScore;
    }
    if (b.score.budgetScore !== a.score.budgetScore) {
      return b.score.budgetScore - a.score.budgetScore;
    }
    return a.price - b.price;
  });
  renderTop3Cards(topRecommendations.slice(0, 3));

  // Render Full Catalog Table (sorted according to state.filters.sortBy)
  renderCatalogTable(filteredEligible);

  // Render Disqualified Section
  renderDisqualifiedSection(filteredDisqualified);

  // Render Compare Floating Dock
  renderCompareDock();

  // Update Brand filter dropdown
  updateBrandFilterOptions();
}

function renderMetricsBar(eligibleList, disqualifiedList) {
  const topScoreEl = document.getElementById("metric-top-score");
  const eligibleCountEl = document.getElementById("metric-eligible-count");
  const disqualifiedCountEl = document.getElementById("metric-disqualified-count");
  const avgPriceEl = document.getElementById("metric-avg-price");

  if (eligibleList.length > 0) {
    const topPct = Math.round(Math.max(...eligibleList.map(m => m.score.totalScore)));
    topScoreEl.textContent = `${topPct}%`;
    topScoreEl.className = "text-xl font-bold " + (topPct >= 80 ? "text-emerald-700" : (topPct >= 50 ? "text-amber-700" : "text-red-600"));
    const avgPrice = Math.round(eligibleList.reduce((acc, m) => acc + m.price, 0) / eligibleList.length);
    avgPriceEl.textContent = `฿${avgPrice.toLocaleString()}`;
  } else {
    topScoreEl.textContent = "0%";
    topScoreEl.className = "text-xl font-bold text-red-600";
    avgPriceEl.textContent = "—";
  }

  if (eligibleCountEl) eligibleCountEl.textContent = eligibleList.length;
  if (disqualifiedCountEl) disqualifiedCountEl.textContent = disqualifiedList.length;

  const disqHeaderBadge = document.getElementById("disqualified-header-count");
  if (disqHeaderBadge) disqHeaderBadge.textContent = `${disqualifiedList.length} Models`;
}

/**
 * Renders the "Top 3 Recommended Machines" Cards (Rank #1, #2, #3)
 * with Human-Readable Match Rate Percentages and Dynamic Thai Explanations
 */
function renderTop3Cards(top3) {
  const container = document.getElementById("top-3-container");
  if (!container) return;

  if (!top3 || top3.length === 0) {
    container.innerHTML = `
      <div class="col-span-1 md:col-span-3 card-coffee p-8 text-center bg-amber-50/60 border-dashed border-2 border-amber-300 rounded-2xl">
        <div class="text-3xl mb-2">☕</div>
        <h4 class="text-base font-bold text-stone-900">${state.lang === 'th' ? 'ไม่พบเครื่องที่ผ่านเกณฑ์' : 'No Matching Machines Found'}</h4>
        <p class="text-xs text-stone-600 mt-1 max-w-md mx-auto">
          ${state.lang === 'th' ? 'ลองเพิ่มงบประมาณ ลดจำนวนแก้วที่ต้องการ หรือเลือกระบบไฟฟ้าแรงสูง' : 'Try increasing your budget, lowering required cups, or enabling Heavy power.'}
        </p>
      </div>
    `;
    return;
  }

  const badgeStyles = ["top3-badge-1", "top3-badge-2", "top3-badge-3"];

  container.innerHTML = top3.map((m, index) => {
    const s = m.score;
    const h = getHumanScoring(m, state.user, s);
    const isCompared = state.comparedModelIds.includes(m.id);
    const badgeClass = badgeStyles[index] || "bg-stone-800 text-white";

    return `
      <div class="card-coffee top3-card p-5 flex flex-col justify-between border ${index === 0 ? 'border-amber-400 ring-2 ring-amber-400/20 shadow-md' : 'border-stone-200'}">
        <div>
          <!-- Header: Rank Badge & Prominent Match Rate Badge -->
          <div class="flex items-start justify-between gap-2 mb-3">
            <div>
              ${index === 0
                ? `<span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-xl text-xs font-black top3-badge-1 shadow-sm">
                     <span>🏆 #1 ${state.lang === 'th' ? 'แนะนำสูงสุด' : 'Top Match'}</span>
                   </span>`
                : `<span class="inline-flex items-center justify-center w-8 h-8 rounded-xl text-xs font-black ${badgeClass}">
                     #${index + 1}
                   </span>`
              }
            </div>

            <div class="text-right">
              <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-xl text-xs font-black border ${h.overall.badgeClass} shadow-sm">
                <span class="w-2 h-2 rounded-full" style="background-color: ${h.overall.color};"></span>
                <span>${h.overall.pct}% ${state.lang === 'th' ? 'เหมาะสมกับคุณ' : 'Match Rate'}</span>
              </span>
            </div>
          </div>

          <!-- Product Identity -->
          <div class="mb-3">
            <span class="text-[11px] font-bold uppercase tracking-wider text-amber-800 block">${m.brand}</span>
            <h4 class="text-base font-serif font-bold text-stone-900 leading-tight">${m.model}</h4>
            <div class="mt-1 flex items-baseline gap-2">
              <span class="text-lg font-extrabold text-stone-900">฿${m.price.toLocaleString()}</span>
              <span class="text-[11px] text-stone-500">(${Math.round(s.priceRatio * 100)}% ${state.lang === 'th' ? 'ของงบประมาณ' : 'of budget'})</span>
            </div>
          </div>

          <!-- Explainability Text -->
          <div class="p-2.5 rounded-xl bg-stone-50 border border-stone-200/80 mb-3 text-[11px] text-stone-700 leading-snug">
            💡 <strong>${state.lang === 'th' ? 'จุดเด่น:' : 'Insight:'}</strong> ${s.explainability}
          </div>

          <!-- 5 Dimensional Breakdown Bars (Match Rate & Thai Micro-Copy) -->
          <div class="space-y-2 py-2.5 border-t border-stone-100 text-xs">
            <!-- 1. Budget -->
            <div class="dimension-row">
              <div class="flex justify-between items-center text-xs mb-1">
                <span class="font-semibold text-stone-700">${h.budget.title}</span>
                <span class="font-bold text-xs" style="color: ${h.budget.color};">${h.budget.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.budget.pct}%; background-color: ${h.budget.color};"></div>
              </div>
              <p class="dimension-note mt-1">${h.budget.note}</p>
            </div>

            <!-- 2. Capacity -->
            <div class="dimension-row">
              <div class="flex justify-between items-center text-xs mb-1">
                <span class="font-semibold text-stone-700">${h.capacity.title}</span>
                <span class="font-bold text-xs" style="color: ${h.capacity.color};">${h.capacity.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.capacity.pct}%; background-color: ${h.capacity.color};"></div>
              </div>
              <p class="dimension-note mt-1">${h.capacity.note}</p>
            </div>

            <!-- 3. Usage -->
            <div class="dimension-row">
              <div class="flex justify-between items-center text-xs mb-1">
                <span class="font-semibold text-stone-700">${h.usage.title}</span>
                <span class="font-bold text-xs" style="color: ${h.usage.color};">${h.usage.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.usage.pct}%; background-color: ${h.usage.color};"></div>
              </div>
              <p class="dimension-note mt-1">${h.usage.note}</p>
            </div>

            <!-- 4. Power -->
            <div class="dimension-row">
              <div class="flex justify-between items-center text-xs mb-1">
                <span class="font-semibold text-stone-700">${h.power.title}</span>
                <span class="font-bold text-xs" style="color: ${h.power.color};">${h.power.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.power.pct}%; background-color: ${h.power.color};"></div>
              </div>
              <p class="dimension-note mt-1">${h.power.note}</p>
            </div>

            <!-- 5. Features -->
            <div class="dimension-row">
              <div class="flex justify-between items-center text-xs mb-1">
                <span class="font-semibold text-stone-700">${h.features.title}</span>
                <span class="font-bold text-xs" style="color: ${h.features.color};">${h.features.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.features.pct}%; background-color: ${h.features.color};"></div>
              </div>
              <div class="flex flex-wrap gap-1 mt-1.5">
                ${h.features.badges.map(b => `
                  <span class="px-2 py-0.5 rounded text-[10px] font-semibold ${b.active ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-stone-100 text-stone-500 border border-stone-200'}">
                    ${b.label}
                  </span>
                `).join('')}
              </div>
            </div>
          </div>
        </div>

        <!-- Card Actions -->
        <div class="pt-3 border-t border-stone-100 flex items-center gap-2">
          <button onclick="openDetailModal('${m.id}')" class="flex-1 py-1.5 px-2.5 bg-stone-100 hover:bg-stone-200 text-stone-800 font-semibold rounded-lg text-xs transition">
            ${state.lang === 'th' ? 'ดูสเปก' : 'Details'}
          </button>
          <button onclick="toggleCompare('${m.id}')" class="py-1.5 px-3 rounded-lg text-xs font-semibold transition ${isCompared ? 'bg-amber-600 text-white' : 'bg-white border border-stone-300 hover:bg-stone-50 text-stone-700'}">
            ${isCompared ? (state.lang === 'th' ? '✓ เลือกแล้ว' : '✓ Added') : (state.lang === 'th' ? '+ เทียบ' : '+ Compare')}
          </button>
        </div>
      </div>
    `;
  }).join("");
}

/**
 * Renders the Full Catalog Table of all passing products sorted by Match Rate descending
 */
function renderCatalogTable(eligibleList) {
  const tbody = document.getElementById("catalog-table-body");
  const noResultsState = document.getElementById("no-results-state");
  if (!tbody) return;

  if (!eligibleList || eligibleList.length === 0) {
    tbody.innerHTML = "";
    if (noResultsState) noResultsState.classList.remove("hidden");
    return;
  }

  if (noResultsState) noResultsState.classList.add("hidden");

  tbody.innerHTML = eligibleList.map((m, idx) => {
    const s = m.score;
    const h = getHumanScoring(m, state.user, s);
    const isCompared = state.comparedModelIds.includes(m.id);
    const loadPct = Math.round(s.ratio * 100);

    return `
      <tr class="hover:bg-amber-50/40 transition">
        <td class="text-center font-bold ${idx < 3 ? 'text-amber-800' : 'text-stone-500'}">
          ${idx + 1}
        </td>
        <td>
          <div class="font-bold text-stone-900">${m.model}</div>
          <div class="text-[11px] text-stone-500">${m.brand} • <span class="font-mono text-[10px]">${m.id}</span></div>
        </td>
        <td>
          <span class="px-2 py-0.5 rounded text-[11px] font-semibold bg-stone-100 text-stone-700">
            ${m.segment}
          </span>
        </td>
        <td>
          <span class="px-2 py-0.5 rounded text-[11px] font-semibold ${m.power === 'Heavy' ? 'bg-orange-100 text-orange-800' : 'bg-stone-100 text-stone-600'}">
            ${m.power}
          </span>
        </td>
        <td class="text-xs">
          <div class="flex flex-wrap gap-1">
            <span class="px-1.5 py-0.5 rounded text-[10px] ${m.dualBoiler ? 'bg-amber-100 text-amber-900 font-bold' : 'bg-stone-100 text-stone-500'}">
              ${m.dualBoiler ? 'DB' : 'Single/HX'}
            </span>
            <span class="px-1.5 py-0.5 rounded text-[10px] ${m.directWater ? 'bg-blue-100 text-blue-900 font-bold' : 'bg-stone-100 text-stone-500'}">
              ${m.directWater ? 'Plumbed' : 'Tank'}
            </span>
          </div>
        </td>
        <td class="text-right font-bold text-stone-900">
          ฿${m.price.toLocaleString()}
        </td>
        <td class="text-right">
          <div class="font-semibold text-stone-800">${m.maxCups} cups</div>
          <div class="text-[11px] text-stone-400">${loadPct}% load</div>
        </td>
        <td class="text-center">
          <div class="inline-flex flex-col items-center gap-1">
            <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-black border ${h.overall.badgeClass}">
              ${h.overall.pct}%
            </span>
            <div class="w-14 h-1.5 bg-stone-100 rounded-full overflow-hidden">
              <div class="h-full rounded-full transition-all duration-500" style="width: ${h.overall.pct}%; background-color: ${h.overall.color};"></div>
            </div>
          </div>
        </td>
        <td class="text-center">
          <div class="inline-flex items-center gap-1">
            <button onclick="openDetailModal('${m.id}')" class="p-1.5 text-stone-600 hover:text-stone-900 hover:bg-stone-100 rounded-lg text-xs" title="View Details">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg>
            </button>
            <button onclick="toggleCompare('${m.id}')" class="p-1.5 rounded-lg text-xs transition ${isCompared ? 'bg-amber-600 text-white' : 'text-stone-600 hover:bg-stone-100'}" title="${isCompared ? 'Remove Compare' : 'Add to Compare'}">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/></svg>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join("");
}

function renderDisqualifiedSection(disqualifiedList) {
  const container = document.getElementById("disqualified-list-container");
  if (!container) return;

  if (disqualifiedList.length === 0) {
    container.innerHTML = `
      <div class="p-4 text-center text-xs text-stone-500 bg-white rounded-xl border border-stone-200">
        ${state.lang === 'th' ? 'เครื่องทุกรุ่นในไฟล์ผ่านเกณฑ์ขั้นต้น! ไม่มีรุ่นใดถูกตัดออก' : 'All machines meet your pre-filter constraints! Zero disqualified.'}
      </div>
    `;
    return;
  }

  container.innerHTML = disqualifiedList.map(m => {
    return `
      <div class="disqualified-card flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div class="flex-1">
          <div class="flex items-center gap-2 mb-1">
            <span class="text-[11px] font-bold px-2 py-0.5 rounded bg-red-100 text-red-800 border border-red-200">
              ${state.lang === 'th' ? 'ไม่ผ่านเกณฑ์ (0 คะแนน)' : 'DISQUALIFIED (0 pts)'}
            </span>
            <span class="text-xs text-stone-500">${m.segment} • ${m.power}</span>
          </div>

          <h4 class="text-sm font-bold text-stone-800">
            <span class="text-stone-500 font-normal">${m.brand}</span> ${m.model}
          </h4>
          <div class="text-xs text-stone-500 mt-0.5">
            ฿${m.price.toLocaleString()} • Max: ${m.maxCups} cups/day
          </div>

          <div class="mt-2 flex flex-wrap gap-1.5">
            ${m.violations.map(v => `
              <span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium bg-red-50 text-red-700 border border-red-200">
                <svg class="w-3.5 h-3.5 text-red-500 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clip-rule="evenodd"/></svg>
                <span><strong>${v.title}:</strong> ${v.description}</span>
              </span>
            `).join("")}
          </div>
        </div>

        <button onclick="openDetailModal('${m.id}')" class="self-start md:self-center px-3 py-1.5 rounded-lg text-xs font-medium bg-white hover:bg-stone-100 text-stone-700 border border-stone-300 transition">
          ${state.lang === 'th' ? 'ดูข้อมูล' : 'Details'}
        </button>
      </div>
    `;
  }).join("");
}

function renderCompareDock() {
  const dock = document.getElementById("compare-dock");
  const countEl = document.getElementById("compare-count-text");
  const previewEl = document.getElementById("compare-items-preview");
  if (!dock) return;

  const count = state.comparedModelIds.length;
  if (count > 0) {
    dock.classList.add("visible");
    countEl.textContent = `${count} of 3 selected`;

    previewEl.innerHTML = state.comparedModelIds.map(id => {
      const machine = catalog.find(m => m.id === id);
      if (!machine) return "";
      return `
        <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs bg-stone-800 text-amber-200 border border-stone-700">
          <span>${machine.brand} ${machine.model.split(" ")[0]}</span>
          <button onclick="toggleCompare('${id}')" class="text-stone-400 hover:text-white font-bold text-xs">✕</button>
        </span>
      `;
    }).join("");
  } else {
    dock.classList.remove("visible");
  }
}

function updateBrandFilterOptions() {
  const select = document.getElementById("filter-brand-select");
  if (!select) return;

  const brands = Array.from(new Set(catalog.map(m => m.brand))).filter(Boolean).sort();
  const currentVal = select.value;

  select.innerHTML = `<option value="all">${state.lang === 'th' ? 'ทุกแบรนด์' : 'All Brands'} (${brands.length})</option>` +
    brands.map(b => `<option value="${b}" ${currentVal === b ? 'selected' : ''}>${b}</option>`).join("");
}

// 9. COMPARISON & DETAIL MODALS

function toggleCompare(machineId) {
  const index = state.comparedModelIds.indexOf(machineId);
  if (index > -1) {
    state.comparedModelIds.splice(index, 1);
  } else {
    if (state.comparedModelIds.length >= 3) {
      alert(state.lang === 'th' ? "เปรียบเทียบได้สูงสุดครั้งละ 3 เครื่องครับ" : "You can compare up to 3 machines simultaneously.");
      return;
    }
    state.comparedModelIds.push(machineId);
  }
  renderUI();
}

function clearComparison() {
  state.comparedModelIds = [];
  renderUI();
}

function openComparisonModal() {
  if (state.comparedModelIds.length === 0) return;

  const modal = document.getElementById("comparison-modal");
  const content = document.getElementById("comparison-modal-content");
  if (!modal || !content) return;

  const machines = state.comparedModelIds.map(id => {
    return state.evaluatedCatalog.find(m => m.id === id);
  }).filter(Boolean);

  content.innerHTML = `
    <div class="p-6 md:p-8">
      <div class="flex items-center justify-between pb-4 border-b border-stone-200 mb-6">
        <div>
          <span class="text-xs font-bold text-amber-800 uppercase tracking-wider">${state.lang === 'th' ? 'การประเมินแบบเคียงข้างกัน' : 'Side-by-Side Evaluation'}</span>
          <h2 class="text-2xl font-serif font-bold text-stone-900">${state.lang === 'th' ? 'ตารางเปรียบเทียบคุณสมบัติเครื่องชงกาแฟ' : 'Machine Comparison Matrix'}</h2>
        </div>
        <button onclick="closeModal('comparison-modal')" class="w-9 h-9 rounded-full bg-stone-100 hover:bg-stone-200 text-stone-600 flex items-center justify-center font-bold text-lg">
          ✕
        </button>
      </div>

      <div class="overflow-x-auto">
        <table class="compare-table">
          <thead>
            <tr>
              <th class="w-1/4">${state.lang === 'th' ? 'คุณสมบัติการประเมิน' : 'Evaluation Attribute'}</th>
              ${machines.map(m => `
                <th class="w-1/4">
                  <div class="text-xs text-amber-800 uppercase font-bold">${m.brand}</div>
                  <div class="text-base font-serif font-bold text-stone-900">${m.model}</div>
                  <div class="mt-2 text-xl font-black ${m.isEligible ? (m.score.totalScore >= 75 ? 'text-emerald-700' : 'text-amber-700') : 'text-red-600'}">
                    ${m.isEligible ? `${m.score.totalScore}/100` : (state.lang === 'th' ? 'ไม่ผ่านเกณฑ์' : 'Disqualified')}
                  </div>
                </th>
              `).join("")}
            </tr>
          </thead>
          <tbody>
            <tr>
              <td class="font-semibold text-stone-700">${state.lang === 'th' ? 'ราคาจำหน่าย' : 'Retail Price'}</td>
              ${machines.map(m => `
                <td>
                  <span class="text-base font-bold text-stone-900">฿${m.price.toLocaleString()}</span>
                  <div class="text-xs text-stone-500">${Math.round(m.price / state.user.maxBudget * 100)}% of budget</div>
                </td>
              `).join("")}
            </tr>
            <tr>
              <td class="font-semibold text-stone-700">${state.lang === 'th' ? 'กำลังผลิตแก้วสูงสุด/วัน' : 'Daily Cup Capacity'}</td>
              ${machines.map(m => `
                <td>
                  <span class="font-bold text-stone-900">${m.maxCups} cups/day</span>
                  <div class="text-xs text-stone-500">${Math.round(state.user.requiredCups / m.maxCups * 100)}% load</div>
                </td>
              `).join("")}
            </tr>
            <tr>
              <td class="font-semibold text-stone-700">${state.lang === 'th' ? 'กลุ่มการใช้งาน' : 'Segment'}</td>
              ${machines.map(m => `
                <td>
                  <span class="px-2.5 py-1 rounded-md text-xs font-semibold ${m.segment === state.user.segment ? 'bg-emerald-100 text-emerald-800' : 'bg-stone-100 text-stone-700'}">
                    ${m.segment}
                  </span>
                </td>
              `).join("")}
            </tr>
            <tr>
              <td class="font-semibold text-stone-700">${state.lang === 'th' ? 'ระบบไฟฟ้า' : 'Electrical Power'}</td>
              ${machines.map(m => `
                <td>
                  <span class="px-2.5 py-1 rounded-md text-xs font-semibold ${m.power === 'Heavy' ? 'bg-orange-100 text-orange-800' : 'bg-stone-100 text-stone-700'}">
                    ${m.power}
                  </span>
                </td>
              `).join("")}
            </tr>
            <tr>
              <td class="font-semibold text-stone-700">${state.lang === 'th' ? 'หม้อต้มคู่ (Dual Boiler)' : 'Dual Boiler System'}</td>
              ${machines.map(m => `
                <td>
                  <span class="font-medium text-xs ${m.dualBoiler ? 'text-emerald-700 font-bold' : 'text-stone-500'}">
                    ${m.dualBoiler ? '✓ Yes (Dual Boiler)' : '✗ Single / HX'}
                  </span>
                </td>
              `).join("")}
            </tr>
            <tr>
              <td class="font-semibold text-stone-700">${state.lang === 'th' ? 'การต่อท่อน้ำตรง' : 'Direct Water Plumbing'}</td>
              ${machines.map(m => `
                <td>
                  <span class="font-medium text-xs ${m.directWater ? 'text-blue-700 font-bold' : 'text-stone-500'}">
                    ${m.directWater ? '✓ Yes (Plumbed)' : '✗ Reservoir Tank'}
                  </span>
                </td>
              `).join("")}
            </tr>
            <!-- Breakdown Rows -->
            <tr class="bg-amber-50/50">
              <td colspan="${machines.length + 1}" class="font-bold text-xs uppercase tracking-wider text-amber-900 py-3">
                ${state.lang === 'th' ? 'ความเหมาะสมตามมิติ (Match Rate %)' : 'Dimension Match Rates (%)'}
              </td>
            </tr>
            <tr>
              <td class="text-xs font-semibold text-stone-700">1. Budget Match</td>
              ${machines.map(m => {
                const pct = Math.round((m.score.budgetScore / 30) * 100);
                return `<td><span class="font-bold text-sm ${pct >= 80 ? 'text-emerald-700' : (pct >= 50 ? 'text-amber-700' : 'text-red-600')}">${pct}%</span></td>`;
              }).join("")}
            </tr>
            <tr>
              <td class="text-xs font-semibold text-stone-700">2. Capacity Match</td>
              ${machines.map(m => {
                const pct = Math.round((m.score.capacityScore / 25) * 100);
                return `<td><span class="font-bold text-sm ${pct >= 80 ? 'text-emerald-700' : (pct >= 50 ? 'text-amber-700' : 'text-red-600')}">${pct}%</span></td>`;
              }).join("")}
            </tr>
            <tr>
              <td class="text-xs font-semibold text-stone-700">3. Usage Match</td>
              ${machines.map(m => {
                const pct = Math.round((m.score.usageScore / 20) * 100);
                return `<td><span class="font-bold text-sm ${pct >= 80 ? 'text-emerald-700' : (pct >= 50 ? 'text-amber-700' : 'text-red-600')}">${pct}%</span></td>`;
              }).join("")}
            </tr>
            <tr>
              <td class="text-xs font-semibold text-stone-700">4. Power Match</td>
              ${machines.map(m => {
                const pct = Math.round((m.score.powerScore / 10) * 100);
                return `<td><span class="font-bold text-sm ${pct >= 80 ? 'text-emerald-700' : (pct >= 50 ? 'text-amber-700' : 'text-red-600')}">${pct}%</span></td>`;
              }).join("")}
            </tr>
            <tr>
              <td class="text-xs font-semibold text-stone-700">5. Feature Fit</td>
              ${machines.map(m => {
                const pct = Math.round((m.score.featureScore / 15) * 100);
                return `<td><span class="font-bold text-sm ${pct >= 80 ? 'text-emerald-700' : (pct >= 50 ? 'text-amber-700' : 'text-red-600')}">${pct}%</span></td>`;
              }).join("")}
            </tr>
            <tr class="bg-stone-100 font-bold">
              <td class="text-sm font-serif">${state.lang === 'th' ? 'ความเหมาะสมโดยรวม' : 'Overall Match Rate'}</td>
              ${machines.map(m => {
                const totalPct = Math.round(m.score.totalScore);
                return `
                <td>
                  <span class="text-xl font-black ${totalPct >= 80 ? 'text-emerald-700' : (totalPct >= 50 ? 'text-amber-700' : 'text-red-600')}">
                    ${totalPct}%
                  </span>
                </td>
                `;
              }).join("")}
            </tr>
          </tbody>
        </table>
      </div>

      <div class="mt-6 flex justify-end gap-3">
        <button onclick="clearComparison(); closeModal('comparison-modal');" class="px-4 py-2 text-stone-600 hover:text-stone-900 text-xs font-semibold">
          ${state.lang === 'th' ? 'ล้างการเปรียบเทียบ' : 'Clear Comparison'}
        </button>
        <button onclick="window.print()" class="px-4 py-2 bg-stone-900 text-white rounded-xl text-xs font-bold hover:bg-stone-800 transition">
          ${state.lang === 'th' ? 'พิมพ์รายงาน' : 'Print Report'}
        </button>
      </div>
    </div>
  `;

  modal.classList.add("open");
}

function openDetailModal(machineId) {
  const machine = state.evaluatedCatalog.find(m => m.id === machineId) || catalog.find(m => m.id === machineId);
  if (!machine) return;

  const modal = document.getElementById("detail-modal");
  const content = document.getElementById("detail-modal-content");
  if (!modal || !content) return;

  const s = machine.score || calculateMCDM(machine, state.user, true);
  const h = getHumanScoring(machine, state.user, s);
  const isEligible = machine.isEligible !== false;
  const isCompared = state.comparedModelIds.includes(machine.id);

  content.innerHTML = `
    <div class="p-6 md:p-8">
      <div class="flex items-start justify-between pb-5 border-b border-stone-200">
        <div>
          <div class="flex items-center gap-2 mb-1.5">
            <span class="text-xs font-bold uppercase tracking-wider text-amber-800">${machine.brand}</span>
            <span class="px-2 py-0.5 rounded text-xs font-semibold bg-stone-100 text-stone-700">${machine.segment}</span>
            <span class="px-2 py-0.5 rounded text-xs font-semibold ${machine.power === 'Heavy' ? 'bg-orange-100 text-orange-800' : 'bg-stone-100 text-stone-600'}">${machine.power} Power</span>
          </div>
          <h2 class="text-2xl md:text-3xl font-serif font-bold text-stone-900">${machine.model}</h2>
          <div class="mt-1 text-xs text-stone-500">Hardware ID: ${machine.id}</div>
        </div>

        <button onclick="closeModal('detail-modal')" class="w-9 h-9 rounded-full bg-stone-100 hover:bg-stone-200 text-stone-600 flex items-center justify-center font-bold text-lg">
          ✕
        </button>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-3 gap-6 py-6">
        <div class="md:col-span-2 space-y-6">
          <div class="p-4 rounded-xl border ${isEligible ? 'bg-amber-50/70 border-amber-200' : 'bg-red-50/70 border-red-200'} flex items-center justify-between">
            <div>
              <div class="text-xs font-bold uppercase tracking-wider ${isEligible ? 'text-amber-800' : 'text-red-800'}">
                ${isEligible ? (state.lang === 'th' ? 'ระดับความเหมาะสมกับคุณ' : 'Your Compatibility Match Rate') : (state.lang === 'th' ? 'สถานะไม่ผ่านเกณฑ์' : 'Disqualification Status')}
              </div>
              <div class="text-2xl font-black ${isEligible ? 'text-stone-900' : 'text-red-700'} mt-0.5">
                ${isEligible ? `${h.overall.pct}% ${state.lang === 'th' ? 'เหมาะสมกับคุณ' : 'Match Rate'}` : (state.lang === 'th' ? 'ไม่ผ่านเกณฑ์ (0%)' : '0% (Disqualified)')}
              </div>
            </div>

            <div class="score-badge-circle border-2 ${getScoreColorClass(h.overall.pct)}">
              <span class="text-xl font-black">${h.overall.pct}%</span>
            </div>
          </div>

          <div class="space-y-3">
            <h4 class="text-sm font-bold uppercase tracking-wider text-stone-800">${state.lang === 'th' ? 'การประเมินแยกตามมิติ (Match Rate)' : 'Dimensional Match Rates'}</h4>
            
            <div class="p-3 bg-stone-50 rounded-xl border border-stone-200 text-xs">
              <div class="flex justify-between font-bold mb-1">
                <span>${h.budget.title}</span>
                <span class="font-bold text-xs" style="color: ${h.budget.color};">${h.budget.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.budget.pct}%; background-color: ${h.budget.color};"></div>
              </div>
              <p class="dimension-note mt-1.5">${h.budget.note}</p>
            </div>

            <div class="p-3 bg-stone-50 rounded-xl border border-stone-200 text-xs">
              <div class="flex justify-between font-bold mb-1">
                <span>${h.capacity.title}</span>
                <span class="font-bold text-xs" style="color: ${h.capacity.color};">${h.capacity.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.capacity.pct}%; background-color: ${h.capacity.color};"></div>
              </div>
              <p class="dimension-note mt-1.5">${h.capacity.note}</p>
            </div>

            <div class="p-3 bg-stone-50 rounded-xl border border-stone-200 text-xs">
              <div class="flex justify-between font-bold mb-1">
                <span>${h.usage.title}</span>
                <span class="font-bold text-xs" style="color: ${h.usage.color};">${h.usage.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.usage.pct}%; background-color: ${h.usage.color};"></div>
              </div>
              <p class="dimension-note mt-1.5">${h.usage.note}</p>
            </div>

            <div class="p-3 bg-stone-50 rounded-xl border border-stone-200 text-xs">
              <div class="flex justify-between font-bold mb-1">
                <span>${h.power.title}</span>
                <span class="font-bold text-xs" style="color: ${h.power.color};">${h.power.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.power.pct}%; background-color: ${h.power.color};"></div>
              </div>
              <p class="dimension-note mt-1.5">${h.power.note}</p>
            </div>

            <div class="p-3 bg-stone-50 rounded-xl border border-stone-200 text-xs">
              <div class="flex justify-between font-bold mb-1">
                <span>${h.features.title}</span>
                <span class="font-bold text-xs" style="color: ${h.features.color};">${h.features.pct}%</span>
              </div>
              <div class="dimension-meter">
                <div class="dimension-fill" style="width: ${h.features.pct}%; background-color: ${h.features.color};"></div>
              </div>
              <div class="flex flex-wrap gap-1 mt-2">
                ${h.features.badges.map(b => `
                  <span class="px-2 py-0.5 rounded text-[10px] font-semibold ${b.active ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-stone-100 text-stone-500 border border-stone-200'}">
                    ${b.label}
                  </span>
                `).join('')}
              </div>
            </div>
          </div>
        </div>

        <div class="bg-stone-50 p-5 rounded-2xl border border-stone-200 flex flex-col justify-between">
          <div>
            <h4 class="text-xs font-bold uppercase tracking-wider text-stone-500 mb-3">${state.lang === 'th' ? 'ข้อมูลจำเพาะจากไฟล์ Excel' : 'Specifications from Excel'}</h4>
            <dl class="space-y-3 text-xs">
              <div>
                <dt class="text-stone-400">Retail Price</dt>
                <dd class="text-lg font-bold text-stone-900">฿${machine.price.toLocaleString()}</dd>
              </div>
              <div>
                <dt class="text-stone-400">Daily Capacity</dt>
                <dd class="font-semibold text-stone-800">${machine.maxCups} cups / day</dd>
              </div>
              <div>
                <dt class="text-stone-400">Dual Boiler</dt>
                <dd class="font-semibold text-stone-800">${machine.dualBoiler ? '✓ Yes' : '✗ No'}</dd>
              </div>
              <div>
                <dt class="text-stone-400">Direct Water Line</dt>
                <dd class="font-semibold text-stone-800">${machine.directWater ? '✓ Yes' : '✗ No'}</dd>
              </div>
              <div>
                <dt class="text-stone-400">Power Type</dt>
                <dd class="font-semibold text-stone-800">${machine.power}</dd>
              </div>
            </dl>
          </div>

          <div class="mt-6 pt-4 border-t border-stone-200 space-y-2">
            <button onclick="toggleCompare('${machine.id}'); closeModal('detail-modal');" class="w-full py-2 px-3 ${isCompared ? 'bg-amber-100 text-amber-900 border border-amber-300' : 'bg-stone-900 text-white hover:bg-stone-800'} rounded-xl text-xs font-bold transition">
              ${isCompared ? (state.lang === 'th' ? '✓ นำออกจากการเปรียบเทียบ' : '✓ Remove Compare') : (state.lang === 'th' ? '+ เพิ่มในตารางเปรียบเทียบ' : '+ Add to Compare')}
            </button>
            <button onclick="closeModal('detail-modal')" class="w-full py-1.5 text-stone-600 hover:text-stone-900 text-xs font-semibold">
              ${state.lang === 'th' ? 'ปิดหน้าต่าง' : 'Close'}
            </button>
          </div>
        </div>
      </div>
    </div>
  `;

  modal.classList.add("open");
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.remove("open");
}

// 10. EXPORT & SUMMARY FUNCTIONS

function exportRecommendationJSON() {
  const topRecommendations = [...state.evaluatedCatalog]
    .filter(m => m.isEligible)
    .sort((a, b) => {
      if (b.score.totalScore !== a.score.totalScore) return b.score.totalScore - a.score.totalScore;
      if (b.score.budgetScore !== a.score.budgetScore) return b.score.budgetScore - a.score.budgetScore;
      return a.price - b.price;
    })
    .slice(0, 3);

  const data = {
    exportedAt: new Date().toISOString(),
    sourceFile: "products.xlsx",
    totalCatalogItems: catalog.length,
    userInputCriteria: state.user,
    topRecommendations: topRecommendations,
    allPassingModels: state.evaluatedCatalog.filter(m => m.isEligible)
  };

  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `coffee-recommendations-${Date.now()}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function copyRecommendationSummary() {
  const top3 = [...state.evaluatedCatalog]
    .filter(m => m.isEligible)
    .sort((a, b) => {
      if (b.score.totalScore !== a.score.totalScore) return b.score.totalScore - a.score.totalScore;
      if (b.score.budgetScore !== a.score.budgetScore) return b.score.budgetScore - a.score.budgetScore;
      return a.price - b.price;
    })
    .slice(0, 3);

  if (top3.length === 0) {
    alert(state.lang === 'th' ? "ไม่พบเครื่องที่ผ่านเกณฑ์เพื่อจัดทำรายงาน" : "No passing recommendations found.");
    return;
  }

  const text = `☕ COFFEE MACHINE MCDM RECOMMENDATION REPORT
==================================================
Criteria: Budget ฿${state.user.maxBudget.toLocaleString()} | Cups: ${state.user.requiredCups}/day | Segment: ${state.user.segment} | Power: ${state.user.powerSupport}
Data Source: products.xlsx (${catalog.length} models)

TOP 3 RECOMMENDED MACHINES:
${top3.map((m, idx) => `
#${idx + 1} ${m.brand} ${m.model}
- Total MCDM Score: ${m.score.totalScore}/100 pts
- Price: ฿${m.price.toLocaleString()} (${Math.round(m.score.priceRatio * 100)}% of budget)
- Capacity: ${m.maxCups} cups/day (${Math.round(m.score.ratio * 100)}% load)
- Specs: ${m.dualBoiler ? 'Dual Boiler' : 'Single/HX'}, ${m.directWater ? 'Direct Line' : 'Water Tank'}
- Score Breakdown: Budget ${m.score.budgetScore}/30, Capacity ${m.score.capacityScore}/25, Usage ${m.score.usageScore}/20, Power ${m.score.powerScore}/10, Fit ${m.score.featureScore}/15
`).join("")}
==================================================
Generated by Coffee Machine Recommender (MCDM Engine)`;

  navigator.clipboard.writeText(text).then(() => {
    alert(state.lang === 'th' ? "คัดลอกสรุปผลลง Clipboard สำเร็จ!" : "Summary report copied to clipboard!");
  }).catch(() => {
    alert("Clipboard permission denied.");
  });
}

function toggleLanguage() {
  state.lang = state.lang === "th" ? "en" : "th";
  applyTranslations();
  renderUI();
}

function applyTranslations() {
  if (typeof document === 'undefined') return;
  const dict = I18N[state.lang] || I18N.th;

  const btnText = document.getElementById("lang-text");
  const btnFlag = document.getElementById("lang-flag");
  if (btnText && btnFlag) {
    btnFlag.textContent = state.lang === "th" ? "🇹🇭" : "🇬🇧";
    btnText.textContent = state.lang === "th" ? "ไทย" : "EN";
  }

  document.querySelectorAll("[data-i18n]").forEach(el => {
    const key = el.getAttribute("data-i18n");
    if (dict[key]) {
      el.textContent = dict[key];
    }
  });
}

// 11. EVENT LISTENERS SETUP

function setupEventListeners() {
  // File picker fallback
  const filePicker = document.getElementById("excel-file-picker");
  if (filePicker) {
    filePicker.addEventListener("change", handleManualFileSelection);
  }

  // Budget Slider & Number Input
  const budgetSlider = document.getElementById("input-budget-slider");
  const budgetNumber = document.getElementById("input-budget-number");
  if (budgetSlider && budgetNumber) {
    budgetSlider.addEventListener("input", (e) => {
      const val = parseInt(e.target.value, 10);
      budgetNumber.value = val;
      state.user.maxBudget = val;
      clearPresetHighlight();
      renderUI();
    });
    budgetNumber.addEventListener("change", (e) => {
      let val = parseInt(e.target.value, 10) || 4000;
      val = Math.max(4000, Math.min(400000, val));
      budgetNumber.value = val;
      budgetSlider.value = val;
      state.user.maxBudget = val;
      clearPresetHighlight();
      renderUI();
    });
  }

  // Daily Cups Slider & Number Input
  const cupsSlider = document.getElementById("input-cups-slider");
  const cupsNumber = document.getElementById("input-cups-number");
  if (cupsSlider && cupsNumber) {
    cupsSlider.addEventListener("input", (e) => {
      const val = parseInt(e.target.value, 10);
      cupsNumber.value = val;
      state.user.requiredCups = val;
      clearPresetHighlight();
      renderUI();
    });
    cupsNumber.addEventListener("change", (e) => {
      let val = parseInt(e.target.value, 10) || 5;
      val = Math.max(5, Math.min(500, val));
      cupsNumber.value = val;
      cupsSlider.value = val;
      state.user.requiredCups = val;
      clearPresetHighlight();
      renderUI();
    });
  }

  // Usage Segment Selection
  document.querySelectorAll(".segment-card").forEach(card => {
    card.addEventListener("click", () => {
      const seg = card.getAttribute("data-segment");
      state.user.segment = seg;
      clearPresetHighlight();
      document.querySelectorAll(".segment-card").forEach(c => c.classList.remove("active"));
      card.classList.add("active");
      renderUI();
    });
  });

  // Power Support Toggle
  document.querySelectorAll(".power-toggle-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const pwr = btn.getAttribute("data-power");
      state.user.powerSupport = pwr;
      clearPresetHighlight();
      document.querySelectorAll(".power-toggle-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      renderUI();
    });
  });

  // Dual Boiler Checkbox
  const dbCheckbox = document.getElementById("check-dual-boiler");
  if (dbCheckbox) {
    dbCheckbox.addEventListener("change", (e) => {
      state.user.requireDualBoiler = e.target.checked;
      const pill = e.target.closest(".feature-checkbox");
      if (pill) pill.classList.toggle("checked", e.target.checked);
      clearPresetHighlight();
      renderUI();
    });
  }

  // Direct Water Checkbox
  const dwCheckbox = document.getElementById("check-direct-water");
  if (dwCheckbox) {
    dwCheckbox.addEventListener("change", (e) => {
      state.user.requireDirectWater = e.target.checked;
      const pill = e.target.closest(".feature-checkbox");
      if (pill) pill.classList.toggle("checked", e.target.checked);
      clearPresetHighlight();
      renderUI();
    });
  }

  // Preset Buttons
  document.querySelectorAll(".preset-chip").forEach(chip => {
    chip.addEventListener("click", () => {
      const presetKey = chip.getAttribute("data-preset");
      applyPreset(presetKey);
    });
  });

  // Sort & Filter Dropdowns
  const sortSelect = document.getElementById("sort-select");
  if (sortSelect) {
    sortSelect.addEventListener("change", (e) => {
      state.filters.sortBy = e.target.value;
      sortEvaluatedCatalog();
      renderUI();
    });
  }

  const segmentFilter = document.getElementById("filter-segment-select");
  if (segmentFilter) {
    segmentFilter.addEventListener("change", (e) => {
      state.filters.segment = e.target.value;
      renderUI();
    });
  }

  const brandFilter = document.getElementById("filter-brand-select");
  if (brandFilter) {
    brandFilter.addEventListener("change", (e) => {
      state.filters.brand = e.target.value;
      renderUI();
    });
  }

  // Quick Budget Steppers
  document.querySelectorAll(".budget-step-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const val = parseInt(btn.getAttribute("data-budget"), 10);
      if (budgetSlider && budgetNumber) {
        budgetSlider.value = val;
        budgetNumber.value = val;
        state.user.maxBudget = val;
        clearPresetHighlight();
        renderUI();
      }
    });
  });

  // Modal Backdrop Click Close
  document.querySelectorAll(".modal-overlay").forEach(overlay => {
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) {
        overlay.classList.remove("open");
      }
    });
  });
}

function clearPresetHighlight() {
  document.querySelectorAll(".preset-chip").forEach(c => c.classList.remove("active"));
  state.activePreset = null;
}

function applyPreset(presetKey) {
  const preset = PRESETS[presetKey];
  if (!preset) return;

  state.user = {
    maxBudget: preset.budget,
    requiredCups: preset.requiredCups,
    segment: preset.segment,
    powerSupport: preset.powerSupport,
    requireDualBoiler: preset.requireDualBoiler,
    requireDirectWater: preset.requireDirectWater
  };
  state.activePreset = presetKey;

  const budgetSlider = document.getElementById("input-budget-slider");
  const budgetNumber = document.getElementById("input-budget-number");
  if (budgetSlider) budgetSlider.value = preset.budget;
  if (budgetNumber) budgetNumber.value = preset.budget;

  const cupsSlider = document.getElementById("input-cups-slider");
  const cupsNumber = document.getElementById("input-cups-number");
  if (cupsSlider) cupsSlider.value = preset.requiredCups;
  if (cupsNumber) cupsNumber.value = preset.requiredCups;

  document.querySelectorAll(".segment-card").forEach(c => {
    c.classList.toggle("active", c.getAttribute("data-segment") === preset.segment);
  });

  document.querySelectorAll(".power-toggle-btn").forEach(b => {
    b.classList.toggle("active", b.getAttribute("data-power") === preset.powerSupport);
  });

  const dbCheck = document.getElementById("check-dual-boiler");
  if (dbCheck) {
    dbCheck.checked = preset.requireDualBoiler;
    const pill = dbCheck.closest(".feature-checkbox");
    if (pill) pill.classList.toggle("checked", preset.requireDualBoiler);
  }

  const dwCheck = document.getElementById("check-direct-water");
  if (dwCheck) {
    dwCheck.checked = preset.requireDirectWater;
    const pill = dwCheck.closest(".feature-checkbox");
    if (pill) pill.classList.toggle("checked", preset.requireDirectWater);
  }

  document.querySelectorAll(".preset-chip").forEach(chip => {
    chip.classList.toggle("active", chip.getAttribute("data-preset") === presetKey);
  });

  renderUI();
}

function resetDefaults() {
  applyPreset("prosumer");
}

// 12. BOOTSTRAP APP ON DOM LOAD
if (typeof document !== 'undefined') {
  document.addEventListener("DOMContentLoaded", async () => {
    setupEventListeners();
    await loadCatalogFromExcel();
    applyPreset("prosumer");
  });
}

// Node.js module export for automated test suite
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    normalizeRow,
    loadCatalogFromExcel,
    evaluateConstraints,
    calculateMCDM,
    state,
    PRESETS,
    getCatalog: () => catalog
  };
}
