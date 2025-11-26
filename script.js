const entries = [];
const addBtn = document.getElementById("addBtn");
const addBtnDefaultText = addBtn.textContent;
let editIndex = null;

function terbilang(n) {
  n = Math.floor(Number(n) || 0);
  if (n === 0) return "nol";
  const angka = ["", "satu", "dua", "tiga", "empat", "lima",
    "enam", "tujuh", "delapan", "sembilan"];

  function _tb(x) {
    if (x === 0) return "";
    if (x < 10) return angka[x];
    if (x === 10) return "sepuluh";
    if (x === 11) return "sebelas";
    if (x < 20) return _tb(x - 10) + " belas";
    if (x < 100) {
      const puluh = Math.floor(x / 10);
      const sisa = x % 10;
      return _tb(puluh) + " puluh" + (sisa ? " " + _tb(sisa) : "");
    }
    if (x < 200) return "seratus" + (x > 100 ? " " + _tb(x - 100) : "");
    if (x < 1000) {
      const ratus = Math.floor(x / 100);
      const sisa = x % 100;
      return _tb(ratus) + " ratus" + (sisa ? " " + _tb(sisa) : "");
    }
    if (x < 2000) return "seribu" + (x > 1000 ? " " + _tb(x - 1000) : "");
    if (x < 1_000_000) {
      const ribu = Math.floor(x / 1000);
      const sisa = x % 1000;
      return _tb(ribu) + " ribu" + (sisa ? " " + _tb(sisa) : "");
    }
    if (x < 1_000_000_000) {
      const juta = Math.floor(x / 1_000_000);
      const sisa = x % 1_000_000;
      return _tb(juta) + " juta" + (sisa ? " " + _tb(sisa) : "");
    }
    if (x < 1_000_000_000_000) {
      const miliar = Math.floor(x / 1_000_000_000);
      const sisa = x % 1_000_000_000;
      return _tb(miliar) + " miliar" + (sisa ? " " + _tb(sisa) : "");
    }
    if (x < 1_000_000_000_000_000) {
      const triliun = Math.floor(x / 1_000_000_000_000);
      const sisa = x % 1_000_000_000_000;
      return _tb(triliun) + " triliun" + (sisa ? " " + _tb(sisa) : "");
    }
    return "";
  }
  return _tb(n).trim();
}

function formatBerat2(berat) {
  if (typeof berat !== "number" || isNaN(berat)) return "0,00";
  return berat.toLocaleString("id-ID", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function toTitleCase(str) {
  return String(str || "")
    .toLowerCase()
    .split(" ")
    .filter(Boolean)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}

function segmentsToHtml(segments) {
  return segments.map(seg => {
    const safe = seg.text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    return seg.italic ? "<em>" + safe + "</em>" : safe;
  }).join("");
}

// --------- Builder Perhiasan Umum -----------
function buildRedaksiPerhiasanSegments({
  jumlahWordRaw,
  namaBarang,
  isEmas,
  karat,
  berat,
  jumlahBerlian,
  berlianAsli,
  stones,
  harga
}) {
  const segments = [];
  const jword = (jumlahWordRaw || "satu").toLowerCase();
  const jwordCap = jword.charAt(0).toUpperCase() + jword.slice(1);

  segments.push({ text: jwordCap + " ", italic: false });
  segments.push({ text: namaBarang.trim() + " ", italic: false });

  if (isEmas) {
    segments.push({ text: "ditaksir emas ", italic: false });
    segments.push({ text: String(karat) + " karat ", italic: false });
  } else {
    segments.push({ text: "ditaksir bukan emas ", italic: false });
  }

  const beratStr = formatBerat2(berat);
  segments.push({ text: "berat " + beratStr + " gram", italic: false });

  const matItems = [];

  // Berlian
  if (jumlahBerlian && jumlahBerlian > 0) {
    const jwNum = String(jumlahBerlian);
    const jwWords = terbilang(jumlahBerlian);
    const jenis = berlianAsli ? "berlian" : "bukan berlian";

    matItems.push([
      { text: " " + jwNum + " ", italic: false },
      { text: "(" + jwWords + ")", italic: true },
      { text: " butir ", italic: false },
      { text: jenis, italic: false }
    ]);
  }

  // Batu lain
  (stones || []).forEach(stone => {
    const j = stone.jumlah || 0;
    const jenis = (stone.jenis || "").trim();
    if (j > 0 && jenis) {
      const jwNum = String(j);
      const jwWords = terbilang(j);
      matItems.push([
        { text: " " + jwNum + " ", italic: false },
        { text: "(" + jwWords + ")", italic: true },
        { text: " butir ", italic: false },
        { text: toTitleCase(jenis), italic: true }
      ]);
    }
  });

  if (matItems.length > 0) {
    segments.push({ text: " bermatakan", italic: false });
    matItems.forEach((itemSeg, idx) => {
      if (idx === 0) {
        segments.push(...itemSeg);
      } else if (idx === matItems.length - 1) {
        segments.push({ text: ", dan", italic: false });
        segments.push(...itemSeg);
      } else {
        segments.push({ text: ", ", italic: false });
        segments.push(...itemSeg);
      }
    });
  }

  harga = Number(harga || 0);

  if (harga === 0) {
    segments.push({ text: ", ", italic: false });
    segments.push({ text: "tidak dilakukan penaksiran harga", italic: false });
    segments.push({ text: ".", italic: false });
    return segments;
  }

  let objekBernilai = 0;
  if (isEmas) objekBernilai++;
  if (jumlahBerlian > 0 && berlianAsli) objekBernilai++;
  if ((stones || []).length > 0) objekBernilai++;

  if (objekBernilai === 0) {
    segments.push({ text: ", ", italic: false });
    segments.push({ text: "tidak dilakukan penaksiran harga", italic: false });
    segments.push({ text: ".", italic: false });
    return segments;
  }

  const hargaFmt = harga.toLocaleString("id-ID");
  const hargaWords = terbilang(harga) + " rupiah";

  segments.push({ text: " ", italic: false });
  segments.push({
    text: objekBernilai === 1
      ? "dengan taksiran harga Rp"
      : "dengan total taksiran harga Rp",
    italic: false
  });
  segments.push({ text: hargaFmt, italic: false });
  segments.push({ text: ",- ", italic: false });
  segments.push({ text: "(", italic: false });
  segments.push({ text: hargaWords, italic: true });
  segments.push({ text: ")", italic: false });
  segments.push({ text: ".", italic: false });

  return segments;
}

// --------- Builder Emas Batangan ----------
function buildRedaksiEmasBatanganSegments({
  jumlahWordRaw,
  merk,
  noSeri,
  isEmas,
  karatEmas,
  berat,
  harga
}) {
  const segments = [];
  const jword = (jumlahWordRaw || "satu").toLowerCase();
  const jwordCap = jword.charAt(0).toUpperCase() + jword.slice(1);

  segments.push({ text: jwordCap + " emas batangan ", italic: false });
  segments.push({ text: "“" + (merk || "").trim() + "” ", italic: false });
  segments.push({ text: "dengan nomor seri " + (noSeri || "").trim() + " ", italic: false });

  if (isEmas) {
    segments.push({ text: "ditaksir emas ", italic: false });
    segments.push({ text: String(karatEmas) + " karat ", italic: false });
  } else {
    segments.push({ text: "ditaksir bukan emas ", italic: false });
  }

  const beratStr = formatBerat2(berat);
  segments.push({ text: "berat " + beratStr + " gram", italic: false });

  harga = Number(harga || 0);
  if (harga === 0) {
    segments.push({ text: ", ", italic: false });
    segments.push({ text: "tidak dilakukan penaksiran harga", italic: false });
    segments.push({ text: ".", italic: false });
    return segments;
  }

  const hargaFmt = harga.toLocaleString("id-ID");
  const hargaWords = terbilang(harga) + " rupiah";

  segments.push({ text: " dengan taksiran harga Rp", italic: false });
  segments.push({ text: hargaFmt, italic: false });
  segments.push({ text: ",- ", italic: false });
  segments.push({ text: "(", italic: false });
  segments.push({ text: hargaWords, italic: true });
  segments.push({ text: ")", italic: false });
  segments.push({ text: ".", italic: false });

  return segments;
}

// --------- Builder Batu Lepasan ----------
function buildRedaksiBatuLepasSegments({
  jumlahWordRaw,
  jenisBatu,
  beratCarat,
  harga
}) {
  const segments = [];
  const jword = (jumlahWordRaw || "satu").toLowerCase();
  const jwordCap = jword.charAt(0).toUpperCase() + jword.slice(1);
  const jenis = (jenisBatu || "").trim();
  const beratStr = formatBerat2(beratCarat);

  segments.push({ text: jwordCap + " ", italic: false });
  segments.push({ text: toTitleCase(jenis) + " ", italic: true });
  segments.push({ text: "berat " + beratStr + " carat", italic: false });

  harga = Number(harga || 0);

  if (harga === 0) {
    segments.push({ text: ", ", italic: false });
    segments.push({ text: "tidak dilakukan penaksiran harga", italic: false });
    segments.push({ text: ".", italic: false });
    return segments;
  }

  const hargaFmt = harga.toLocaleString("id-ID");
  const hargaWords = terbilang(harga) + " rupiah";

  segments.push({ text: " dengan taksiran harga Rp", italic: false });
  segments.push({ text: hargaFmt, italic: false });
  segments.push({ text: ",- ", italic: false });
  segments.push({ text: "(", italic: false });
  segments.push({ text: hargaWords, italic: true });
  segments.push({ text: ")", italic: false });
  segments.push({ text: ".", italic: false });

  return segments;
}

// --------- Render Tabel ----------
function renderTable() {
  const tbody = document.querySelector("#outputTable tbody");
  tbody.innerHTML = "";
  entries.forEach((e, idx) => {
    const tr = document.createElement("tr");

    const tdNo = document.createElement("td");
    tdNo.textContent = e.noUrut;
    tdNo.className = "col-no";

    const tdRed = document.createElement("td");
    tdRed.innerHTML = segmentsToHtml(e.segments);

    const tdAksi = document.createElement("td");
    tdAksi.className = "col-aksi";

    const copyBtn = document.createElement("button");
    copyBtn.textContent = "Salin";
    copyBtn.className = "btn-copy";
    copyBtn.addEventListener("click", () => {
      navigator.clipboard.writeText(e.redaksi).then(() => {
        copyBtn.textContent = "✅ Disalin!";
        setTimeout(() => (copyBtn.textContent = "Salin"), 1200);
      });
    });
    tdAksi.appendChild(copyBtn);

    const editBtn = document.createElement("button");
    editBtn.textContent = "Edit";
    editBtn.className = "btn-edit";
    editBtn.addEventListener("click", () => loadFormFromEntry(idx));
    tdAksi.appendChild(editBtn);

    const delBtn = document.createElement("button");
    delBtn.textContent = "Hapus";
    delBtn.className = "btn-delete";
    delBtn.addEventListener("click", () => {
      entries.splice(idx, 1);
      renderTable();
    });
    tdAksi.appendChild(delBtn);

    tr.appendChild(tdNo);
    tr.appendChild(tdRed);
    tr.appendChild(tdAksi);
    tbody.appendChild(tr);
  });
}

// --------- Tambah / Edit dari Form ----------
function addEntryFromForm() {
  const noUrut = (document.getElementById("noUrut").value || "").trim() || "-";
  const jumlahWordRaw = (document.getElementById("jumlahBarang").value || "satu").trim().toLowerCase();
  const barang = (document.getElementById("barang").value || "").trim().toLowerCase();
  const barangCustom = (document.getElementById("barangCustom").value || "").trim();

  const karatVal = (document.getElementById("karat").value || "").trim();
  const isEmas = (karatVal !== "Diluar SNI");
  const karatInt = isEmas ? parseInt(karatVal || "0", 10) : 0;

  let berat = parseFloat(String(document.getElementById("berat").value).replace(",", "."));
  if (isNaN(berat)) berat = 0;

  const adaBerlian = document.getElementById("adaBerlian").checked;
  let jumlahBerlian = 0;
  let berlianAsli = true;
  if (adaBerlian) {
    jumlahBerlian = parseInt(document.getElementById("jmlBerlian").value || "0", 10);
    const jenisBerlianVal = document.getElementById("jenisBerlian").value;
    berlianAsli = (jenisBerlianVal === "asli");
  }

  const adaBatu = document.getElementById("adaBatu").checked;
  const stones = [];
  if (adaBatu) {
    const b1j = parseInt(document.getElementById("batu1_jml").value || "0", 10);
    const b1t = (document.getElementById("batu1_jenis").value || "").trim();
    if (b1j > 0 && b1t) stones.push({ jumlah: b1j, jenis: b1t });

    const b2j = parseInt(document.getElementById("batu2_jml").value || "0", 10);
    const b2t = (document.getElementById("batu2_jenis").value || "").trim();
    if (b2j > 0 && b2t) stones.push({ jumlah: b2j, jenis: b2t });

    const b3j = parseInt(document.getElementById("batu3_jml").value || "0", 10);
    const b3t = (document.getElementById("batu3_jenis").value || "").trim();
    if (b3j > 0 && b3t) stones.push({ jumlah: b3j, jenis: b3t });
  }

  const merkEmas = (document.getElementById("merkEmas").value || "").trim();
  const noSeriEmas = (document.getElementById("noSeriEmas").value || "").trim();
  const harga = parseInt(document.getElementById("harga").value || "0", 10) || 0;
  const jenisBatuLepas = (document.getElementById("batu1_jenis").value || "").trim();

  // Validasi
  if (barang === "lainnya" && !barangCustom) {
    alert("Untuk opsi 'Lainnya', isi jenis barang pada kolom yang tersedia.");
    return;
  }
  if (berat <= 0) {
    alert("Berat harus lebih besar dari 0.");
    return;
  }
  if (barang === "emas batangan" && (!merkEmas || !noSeriEmas)) {
    alert("Untuk emas batangan, isi Merk Emas dan Nomor Seri terlebih dahulu.");
    return;
  }
  if (barang === "batu" && !jenisBatuLepas) {
    alert("Untuk batu lepasan, isi jenis batu pada kolom Batu 1.");
    return;
  }
  if (adaBerlian && jumlahBerlian <= 0) {
    alert("Jika ada berlian, jumlah berlian harus lebih besar dari 0.");
    return;
  }
  if (adaBatu && stones.length === 0 && barang !== "batu") {
    alert("Checkbox batu dicentang, tetapi belum ada data batu yang diisi.");
    return;
  }

  const namaBarangFinal = (barang === "lainnya" && barangCustom)
    ? barangCustom
    : barang;

  let segments;
  if (barang === "emas batangan") {
    segments = buildRedaksiEmasBatanganSegments({
      jumlahWordRaw,
      merk: merkEmas,
      noSeri: noSeriEmas,
      isEmas,
      karatEmas: karatInt,
      berat,
      harga
    });
  } else if (barang === "batu") {
    segments = buildRedaksiBatuLepasSegments({
      jumlahWordRaw,
      jenisBatu: jenisBatuLepas,
      beratCarat: berat,
      harga
    });
  } else {
    segments = buildRedaksiPerhiasanSegments({
      jumlahWordRaw,
      namaBarang: namaBarangFinal,
      isEmas,
      karat: karatInt,
      berat,
      jumlahBerlian,
      berlianAsli,
      stones,
      harga
    });
  }

  const redaksiText = segments.map(s => s.text).join("");

  const data = {
    noUrut,
    jumlahWordRaw,
    barang,
    barangCustom,
    namaBarangFinal,
    karatVal,
    isEmas,
    karatInt,
    berat,
    adaBerlian,
    jumlahBerlian,
    berlianAsli,
    adaBatu,
    stones,
    merkEmas,
    noSeriEmas,
    harga,
    jenisBatuLepas
  };

  if (editIndex !== null) {
    entries[editIndex] = { noUrut, redaksi: redaksiText, segments, data };
    editIndex = null;
    addBtn.textContent = addBtnDefaultText;
  } else {
    entries.push({ noUrut, redaksi: redaksiText, segments, data });
  }

  renderTable();
}

addBtn.addEventListener("click", addEntryFromForm);

function loadFormFromEntry(index) {
  const entry = entries[index];
  if (!entry || !entry.data) {
    alert("Data asli untuk baris ini tidak tersedia untuk diedit.");
    return;
  }
  const d = entry.data;
  editIndex = index;
  addBtn.textContent = "Simpan Perubahan";

  document.getElementById("noUrut").value = d.noUrut && d.noUrut !== "-" ? d.noUrut : "";
  document.getElementById("jumlahBarang").value = d.jumlahWordRaw || "satu";

  const barangSelectEl = document.getElementById("barang");
  const barangCustomEl = document.getElementById("barangCustom");
  const rowBarangCustomEl = document.getElementById("rowBarangCustom");

  if (d.barang === "lainnya") {
    barangSelectEl.value = "lainnya";
    rowBarangCustomEl.style.display = "flex";
    barangCustomEl.value = d.barangCustom || d.namaBarangFinal || "";
  } else {
    barangSelectEl.value = d.barang || "cincin";
    rowBarangCustomEl.style.display = "none";
    barangCustomEl.value = "";
  }

  const karatEl = document.getElementById("karat");
  if (!d.isEmas) {
    karatEl.value = "Diluar SNI";
  } else {
    karatEl.value = String(d.karatInt || d.karatVal || "24");
  }

  document.getElementById("berat").value =
    typeof d.berat === "number" ? String(d.berat) : "";

  document.getElementById("adaBerlian").checked = !!d.adaBerlian;
  document.getElementById("jmlBerlian").value =
    d.jumlahBerlian != null ? d.jumlahBerlian : "";
  document.getElementById("jenisBerlian").value = d.berlianAsli ? "asli" : "bukan";

  document.getElementById("adaBatu").checked = !!d.adaBatu;
  const stones = d.stones || [];
  const s1 = stones[0] || {};
  const s2 = stones[1] || {};
  const s3 = stones[2] || {};

  document.getElementById("batu1_jml").value = s1.jumlah != null ? s1.jumlah : "";
  document.getElementById("batu1_jenis").value = s1.jenis || "";
  document.getElementById("batu2_jml").value = s2.jumlah != null ? s2.jumlah : "";
  document.getElementById("batu2_jenis").value = s2.jenis || "";
  document.getElementById("batu3_jml").value = s3.jumlah != null ? s3.jumlah : "";
  document.getElementById("batu3_jenis").value = s3.jenis || "";

  if (d.barang === "batu" && d.jenisBatuLepas) {
    document.getElementById("batu1_jenis").value = d.jenisBatuLepas;
  }

  document.getElementById("merkEmas").value = d.merkEmas || "";
  document.getElementById("noSeriEmas").value = d.noSeriEmas || "";
  document.getElementById("harga").value = d.harga != null ? d.harga : "";
}

const barangSelect = document.getElementById("barang");
const rowBarangCustom = document.getElementById("rowBarangCustom");
const barangCustomInput = document.getElementById("barangCustom");
barangSelect.addEventListener("change", () => {
  if (barangSelect.value === "lainnya") {
    rowBarangCustom.style.display = "flex";
  } else {
    rowBarangCustom.style.display = "none";
    barangCustomInput.value = "";
  }
});

// ---------- CSV IMPORT ----------
document.getElementById("importBtn").addEventListener("click", function () {
  const fileInput = document.getElementById("csvInput");
  const file = fileInput.files && fileInput.files[0];
  if (!file) {
    alert("Pilih file CSV terlebih dahulu.");
    return;
  }
  const reader = new FileReader();
  reader.onload = e => importFromCsv(e.target.result);
  reader.readAsText(file, "utf-8");
});

function detectDelimiter(line) {
  const candidates = [",", ";", "\t"];
  let bestDelim = ",";
  let bestCount = 0;
  for (const delim of candidates) {
    const count = line.split(delim).length;
    if (count > bestCount) {
      bestCount = count;
      bestDelim = delim;
    }
  }
  return bestDelim;
}

function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
  if (!lines.length) return [];
  const delim = detectDelimiter(lines[0]);
  return lines.map(line =>
    line.split(delim).map(cell => cell.replace(/^"|"$/g, "").trim())
  );
}

function isTrueLike(val) {
  if (!val) return false;
  const v = String(val).trim().toLowerCase();
  return (v === "1" || v === "true" || v === "ya" || v === "y");
}

function importFromCsv(csvText) {
  const rows = parseCsv(csvText);
  if (rows.length <= 1) {
    alert("File CSV tidak berisi data (minimal 1 baris setelah header).");
    return;
  }

  const header = rows[0].map(h => String(h || "").trim().toLowerCase());

  const idxMap = {
    nourut: header.indexOf("nourut"),
    jumlah: header.indexOf("jumlah"),
    barang: header.indexOf("barang"),
    karat: header.indexOf("karat"),
    berat: header.indexOf("berat"),
    adaberlian: header.indexOf("adaberlian"),
    jmlberlian: header.indexOf("jmlberlian"),
    jenisberlian: header.indexOf("jenisberlian"),
    adabatu: header.indexOf("adabatu"),
    batu1jml: header.indexOf("batu1jml"),
    batu1jenis: header.indexOf("batu1jenis"),
    batu2jml: header.indexOf("batu2jml"),
    batu2jenis: header.indexOf("batu2jenis"),
    batu3jml: header.indexOf("batu3jml"),
    batu3jenis: header.indexOf("batu3jenis"),
    merkemas: header.indexOf("merkemas"),
    noseriemas: header.indexOf("noseriemas"),
    harga: header.indexOf("harga")
  };

  const missingRequired = [];
  if (idxMap.jumlah === -1) missingRequired.push("Jumlah");
  if (idxMap.barang === -1) missingRequired.push("Barang");
  if (idxMap.berat === -1) missingRequired.push("Berat");

  if (missingRequired.length > 0) {
    alert("Header CSV tidak lengkap. Kolom wajib berikut tidak ditemukan: " + missingRequired.join(", "));
    return;
  }

  function getVal(row, key, defaultVal = "") {
    const idx = idxMap[key];
    if (idx === -1 || idx >= row.length) return defaultVal;
    return row[idx];
  }

  let imported = 0;

  for (let i = 1; i < rows.length; i++) {
    const cols = rows[i];
    if (!cols || cols.length === 0) continue;

    let noUrut = (getVal(cols, "nourut", "") || "").trim() || "-";
    let jumlahWordRaw = (getVal(cols, "jumlah", "satu") || "satu").trim().toLowerCase();
    let barang = (getVal(cols, "barang", "cincin") || "cincin").trim().toLowerCase();

    const karatVal = (getVal(cols, "karat", "") || "").trim();
    const isEmas = (karatVal !== "Diluar SNI");
    const karatInt = isEmas ? parseInt(karatVal || "0", 10) : 0;

    const beratStr = getVal(cols, "berat", "0");
    const berat = parseFloat((beratStr || "0").replace(",", ".")) || 0;

    const adaBerlianStr = getVal(cols, "adaberlian", "");
    const jmlBerlianStr = getVal(cols, "jmlberlian", "0");
    const jenisBerlianStr = getVal(cols, "jenisberlian", "");
    const adaBatuStr = getVal(cols, "adabatu", "");
    const b1jStr = getVal(cols, "batu1jml", "0");
    const b1t = getVal(cols, "batu1jenis", "");
    const b2jStr = getVal(cols, "batu2jml", "0");
    const b2t = getVal(cols, "batu2jenis", "");
    const b3jStr = getVal(cols, "batu3jml", "0");
    const b3t = getVal(cols, "batu3jenis", "");
    const merkEmas = getVal(cols, "merkemas", "");
    const noSeriEmas = getVal(cols, "noseriemas", "");
    const hargaStr = getVal(cols, "harga", "0");

    const adaBerlian = isTrueLike(adaBerlianStr);
    let jumlahBerlian = 0;
    let berlianAsli = true;
    if (adaBerlian) {
      jumlahBerlian = parseInt(jmlBerlianStr || "0", 10);
      berlianAsli = ((jenisBerlianStr || "").trim().toLowerCase() !== "bukan");
    }

    const adaBatu = isTrueLike(adaBatuStr);
    const stones = [];
    if (adaBatu) {
      const b1j = parseInt(b1jStr || "0", 10);
      const batu1jenis = (b1t || "").trim();
      if (b1j > 0 && batu1jenis) stones.push({ jumlah: b1j, jenis: batu1jenis });

      const b2j = parseInt(b2jStr || "0", 10);
      const batu2jenis = (b2t || "").trim();
      if (b2j > 0 && batu2jenis) stones.push({ jumlah: b2j, jenis: batu2jenis });

      const b3j = parseInt(b3jStr || "0", 10);
      const batu3jenis = (b3t || "").trim();
      if (b3j > 0 && batu3jenis) stones.push({ jumlah: b3j, jenis: batu3jenis });
    }

    const harga = parseInt(hargaStr || "0", 10) || 0;
    const jenisBatuLepas = (b1t || "").trim();

    let segments;
    if (barang === "emas batangan") {
      segments = buildRedaksiEmasBatanganSegments({
        jumlahWordRaw,
        merk: merkEmas,
        noSeri: noSeriEmas,
        isEmas,
        karatEmas: karatInt,
        berat,
        harga
      });
    } else if (barang === "batu") {
      segments = buildRedaksiBatuLepasSegments({
        jumlahWordRaw,
        jenisBatu: jenisBatuLepas,
        beratCarat: berat,
        harga
      });
    } else {
      const namaBarangFinal = barang;
      segments = buildRedaksiPerhiasanSegments({
        jumlahWordRaw,
        namaBarang: namaBarangFinal,
        isEmas,
        karat: karatInt,
        berat,
        jumlahBerlian,
        berlianAsli,
        stones,
        harga
      });
    }

    const redaksiText = segments.map(s => s.text).join("");

    const data = {
      noUrut,
      jumlahWordRaw,
      barang,
      barangCustom: "",
      namaBarangFinal: barang,
      karatVal,
      isEmas,
      karatInt,
      berat,
      adaBerlian,
      jumlahBerlian,
      berlianAsli,
      adaBatu,
      stones,
      merkEmas,
      noSeriEmas,
      harga,
      jenisBatuLepas: barang === "batu" ? jenisBatuLepas : ""
    };

    entries.push({ noUrut, redaksi: redaksiText, segments, data });
    imported++;
  }

  alert(
    imported === 0
      ? "CSV terbaca, tetapi tidak ada baris valid yang diimpor."
      : "Berhasil mengimpor " + imported + " baris dari CSV."
  );
  renderTable();
}

// ---------- DOWNLOAD TXT ----------
document.getElementById("downloadBtn").addEventListener("click", function () {
  if (entries.length === 0) {
    alert("Belum ada redaksi yang dibuat.");
    return;
  }
  let content = "NoUrut\tRedaksi\r\n";
  entries.forEach(e => {
    const safeRedaksi = e.redaksi.replace(/\s+/g, " ").trim();
    content += e.noUrut + "\t" + safeRedaksi + "\r\n";
  });

  const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "redaksi_taksiran.txt";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});
