// ====== DATA PENYIMPAN ======
const entries = []; // { noUrut, redaksi, segments, data }
const STORAGE_KEY = "redaksiEntries_v2";

// ====== FUNGSI BANTU UMUM ======
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
    const triliun = Math.floor(x / 1_000_000_000_000);
    const sisa = x % 1_000_000_000_000;
    return _tb(triliun) + " triliun" + (sisa ? " " + _tb(sisa) : "");
  }

  return _tb(n).trim();
}

function toTitleCase(str) {
  return (str || "").toLowerCase().split(/\s+/).filter(Boolean).map(w =>
    w.charAt(0).toUpperCase() + w.slice(1)
  ).join(" ");
}

function formatBerat2(berat) {
  const num = Number(berat) || 0;
  return num.toFixed(2).replace(".", ",");
}

function escapeHtml(text) {
  return text.replace(/&/g, "&amp;")
             .replace(/</g, "&lt;")
             .replace(/>/g, "&gt;");
}

function segmentsToHtml(segments) {
  let html = "";
  segments.forEach(seg => {
    const safe = escapeHtml(seg.text).replace(/\n/g, "<br>");
    if (seg.italic) {
      html += "<i>" + safe + "</i>";
    } else {
      html += safe;
    }
  });
  return html;
}

function clearErrors() {
  document.querySelectorAll(".input-error").forEach(el => {
    el.classList.remove("input-error");
  });
  document.querySelectorAll(".error-text").forEach(el => {
    el.style.display = "none";
    el.textContent = "";
  });
}

function setError(inputEl, errorEl, message) {
  if (inputEl) inputEl.classList.add("input-error");
  if (errorEl) {
    errorEl.textContent = message;
    errorEl.style.display = "block";
  }
}

// ====== BUILDER PERHIASAN ======
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
  if (jumlahBerlian > 0) {
    const jw = terbilang(jumlahBerlian);
    matItems.push([
      { text: " " + jumlahBerlian + " ", italic: false },
      { text: "(", italic: false },
      { text: jw, italic: true },        // terbilang italic
      { text: ") butir ", italic: false },
      { text: berlianAsli ? "berlian" : "bukan berlian", italic: false } // "berlian" tidak italic
    ]);
  }

  // Batu: tanpa kata "batu"
  (stones || []).forEach(batu => {
    const jml = Number(batu.jumlah || 0);
    const jenis = (batu.jenis || "").trim();
    if (jml > 0 && jenis) {
      const jw = terbilang(jml);
      matItems.push([
        { text: " " + jml + " ", italic: false },
        { text: "(", italic: false },
        { text: jw, italic: true },     // terbilang italic
        { text: ") butir ", italic: false },
        { text: toTitleCase(jenis), italic: true } // nama batu italic
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
  } else {
    const hargaFmt = harga.toLocaleString("id-ID");
    const hargaWords = terbilang(harga) + " rupiah";

    segments.push({ text: " ", italic: false });
    if (objekBernilai === 1) {
      segments.push({ text: "dengan taksiran harga Rp", italic: false });
    } else {
      segments.push({ text: "dengan total taksiran harga Rp", italic: false });
    }
    segments.push({ text: hargaFmt, italic: false });
    segments.push({ text: ",- ", italic: false });
    segments.push({ text: "(", italic: false });
    segments.push({ text: hargaWords, italic: true }); // terbilang rupiah italic
    segments.push({ text: ")", italic: false });
    segments.push({ text: ".", italic: false });
  }

  return segments;
}

// ====== BUILDER EMAS BATANGAN ======
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

  segments.push({ text: jwordCap + " ", italic: false });
  segments.push({ text: "emas batangan ", italic: false });

  merk = (merk || "").trim();
  noSeri = (noSeri || "").trim();

  if (merk) {
    segments.push({ text: "“", italic: false });
    segments.push({ text: merk, italic: false });
    segments.push({ text: "” ", italic: false });
  }

  if (noSeri) {
    segments.push({ text: "dengan nomor seri ", italic: false });
    segments.push({ text: noSeri + " ", italic: false });
  }

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

// ====== BUILDER BATU LEPASAN ======
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

// ====== RENDER TABEL HASIL ======
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

    // tombol Salin
    const copyBtn = document.createElement("button");
    copyBtn.textContent = "Salin";
    copyBtn.className = "btn-copy";
    copyBtn.addEventListener("click", () => {
      navigator.clipboard.writeText(e.redaksi).then(() => {
        copyBtn.textContent = "✅ Disalin!";
        setTimeout(() => (copyBtn.textContent = "Salin"), 1200);
      }).catch(() => {
        alert("Gagal menyalin ke clipboard. Coba pakai Ctrl+C manual.");
      });
    });
    tdAksi.appendChild(copyBtn);

    // tombol Hapus
    const delBtn = document.createElement("button");
    delBtn.textContent = "Hapus";
    delBtn.className = "btn-delete";
    delBtn.addEventListener("click", () => {
      entries.splice(idx, 1);
      renderTable();
      saveToStorage();
      updateDashboard();
      renderHistory();
    });
    tdAksi.appendChild(delBtn);

    tr.appendChild(tdNo);
    tr.appendChild(tdRed);
    tr.appendChild(tdAksi);
    tbody.appendChild(tr);
  });
}

// ====== LOCALSTORAGE: SIMPAN & MUAT ======
function saveToStorage() {
  const plain = entries.map(e => ({
    noUrut: e.noUrut,
    redaksi: e.redaksi,
    data: e.data || {}
  }));
  localStorage.setItem(STORAGE_KEY, JSON.stringify(plain));
}

function loadFromStorage() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;
  try {
    const arr = JSON.parse(raw);
    arr.forEach(item => {
      const seg = [{ text: item.redaksi, italic: false }];
      entries.push({
        noUrut: item.noUrut,
        redaksi: item.redaksi,
        segments: seg,
        data: item.data || {}
      });
    });
    renderTable();
    updateDashboard();
    renderHistory();
  } catch (e) {
    console.error("Gagal parse storage:", e);
  }
}

// ====== DASHBOARD & RIWAYAT ======
function updateDashboard() {
  const total = entries.length;
  let totalHarga = 0;
  let countEmasBatangan = 0;
  let countBatuLepas = 0;

  entries.forEach(e => {
    const d = e.data || {};
    const h = Number(d.harga || 0);
    if (!isNaN(h)) totalHarga += h;
    const barang = (d.barang || "").toLowerCase();
    if (barang === "emas batangan") countEmasBatangan++;
    if (barang === "batu") countBatuLepas++;
  });

  document.getElementById("dashTotal").textContent = total;
  document.getElementById("dashTotalHarga").textContent = totalHarga.toLocaleString("id-ID");
  document.getElementById("dashEmasBatangan").textContent = countEmasBatangan;
  document.getElementById("dashBatuLepas").textContent = countBatuLepas;
}

function renderHistory() {
  const tbody = document.getElementById("historyTableBody");
  if (!tbody) return;
  tbody.innerHTML = "";
  entries.forEach(e => {
    const d = e.data || {};
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${e.noUrut}</td>
      <td>${d.namaBarangFinal || d.barang || "-"}</td>
      <td style="text-align:right;">${Number(d.harga || 0).toLocaleString("id-ID")}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ====== TAMBAH DARI FORM ======
function addEntryFromForm() {
  clearErrors();

  const noUrut = (document.getElementById("noUrut").value || "").trim() || "-";
  const jumlahWordRaw = document.getElementById("jumlahBarang").value;
  const barang = document.getElementById("barang").value;
  const barangCustom = (document.getElementById("barangCustom").value || "").trim();

  const karatVal = document.getElementById("karat").value;
  const isEmas = (karatVal !== "Diluar SNI");
  const karatInt = isEmas ? parseInt(karatVal, 10) : 0;

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
  const harga = parseInt(document.getElementById("harga").value || "0", 10);
  const jenisBatuLepas = (document.getElementById("batu1_jenis").value || "").trim();

  const beratInput = document.getElementById("berat");
  const hargaInput = document.getElementById("harga");
  const hargaErrorEl = document.getElementById("hargaError");

  // ====== Validasi kontekstual ======
  if (barang === "lainnya" && !barangCustom) {
    alert("Untuk opsi 'Lainnya', isi jenis barang pada kolom yang tersedia.");
    setError(document.getElementById("barangCustom"), null, "");
    return;
  }

  if (berat <= 0) {
    setError(beratInput, null, "Berat harus lebih besar dari 0.");
    alert("Berat harus lebih besar dari 0.");
    return;
  }

  if (barang === "emas batangan" && (!merkEmas || !noSeriEmas)) {
    setError(document.getElementById("merkEmas"), null, "");
    setError(document.getElementById("noSeriEmas"), null, "");
    alert("Untuk emas batangan, isi Merk Emas dan Nomor Seri terlebih dahulu.");
    return;
  }

  if (barang === "batu" && !jenisBatuLepas) {
    setError(document.getElementById("batu1_jenis"), null, "Jenis batu wajib diisi untuk batu lepasan.");
    alert("Untuk batu lepasan, isi jenis batu pada kolom Batu 1.");
    return;
  }

  if (adaBerlian && jumlahBerlian <= 0) {
    setError(document.getElementById("jmlBerlian"), null, "Jumlah berlian harus > 0 jika ada berlian.");
    alert("Jika ada berlian, jumlah berlian harus lebih besar dari 0.");
    return;
  }

  if (adaBatu && stones.length === 0 && barang !== "batu") {
    alert("Checkbox batu dicentang, tetapi belum ada data batu yang diisi.");
    return;
  }

  if (harga === 0 && isEmas) {
    setError(hargaInput, hargaErrorEl,
      "Emas dengan kadar valid biasanya diberi nilai taksiran. Yakin nilai 0?");
    // Tidak di-return, hanya warning visual.
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

  entries.push({ noUrut, redaksi: redaksiText, segments, data });
  renderTable();
  saveToStorage();
  updateDashboard();
  renderHistory();
}

document.getElementById("addBtn").addEventListener("click", addEntryFromForm);

// ====== SHOW/HIDE BARANG CUSTOM UNTUK "LAINNYA" ======
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

// ====== IMPORT CSV ======
document.getElementById("importBtn").addEventListener("click", function () {
  const fileInput = document.getElementById("csvInput");
  const file = fileInput.files && fileInput.files[0];
  if (!file) {
    alert("Pilih file CSV terlebih dahulu.");
    return;
  }
  const reader = new FileReader();
  reader.onload = function (e) {
    const text = e.target.result;
    importFromCsv(text);
  };
  reader.readAsText(file, "utf-8");
});

function detectDelimiter(line) {
  const candidates = [",", ";", "\t"];
  let bestDelim = ",";
  let bestCount = 0;
  candidates.forEach(d => {
    const c = line.split(d).length;
    if (c > bestCount) {
      bestCount = c;
      bestDelim = d;
    }
  });
  return bestDelim;
}

function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
  if (lines.length === 0) return [];
  const delim = detectDelimiter(lines[0]);
  return lines.map(line =>
    line.split(delim).map(cell =>
      cell.replace(/^"|"$/g, "").trim()
    )
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

  let imported = 0;

  // baris 0 = header
  for (let i = 1; i < rows.length; i++) {
    const cols = rows[i];
    if (cols.length < 18) continue;

    let [
      noUrut,
      jumlahWordRaw,
      barang,
      karatVal,
      beratStr,
      adaBerlianStr,
      jmlBerlianStr,
      jenisBerlianStr,
      adaBatuStr,
      b1jStr,
      b1t,
      b2jStr,
      b2t,
      b3jStr,
      b3t,
      merkEmas,
      noSeriEmas,
      hargaStr
    ] = cols;

    noUrut = (noUrut || "").trim() || "-";
    jumlahWordRaw = (jumlahWordRaw || "satu").trim().toLowerCase();
    barang = (barang || "cincin").trim().toLowerCase();

    const isEmas = (karatVal !== "Diluar SNI");
    const karatInt = isEmas ? parseInt(karatVal || "0", 10) : 0;
    const berat = parseFloat((beratStr || "0").replace(",", ".")) || 0;

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

    const harga = parseInt(hargaStr || "0", 10);
    const jenisBatuLepas = (b1t || "").trim();

    let segments;
    let namaBarangFinal = barang;

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
      barangCustom: "",
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
      jenisBatuLepas: barang === "batu" ? jenisBatuLepas : ""
    };

    entries.push({ noUrut, redaksi: redaksiText, segments, data });
    imported++;
  }

  renderTable();
  saveToStorage();
  updateDashboard();
  renderHistory();

  if (imported === 0) {
    alert("CSV terbaca, tetapi tidak ada baris valid yang diimport.\nPeriksa kembali urutan kolom dan delimiter (koma / titik koma).");
  } else {
    alert("Berhasil mengimport " + imported + " baris dari CSV.");
  }
}

// ====== DOWNLOAD TXT ======
document.getElementById("downloadBtn").addEventListener("click", function () {
  if (entries.length === 0) {
    alert("Belum ada redaksi yang dibuat.");
    return;
  }

  let content = "NoUrut\tRedaksi\n";
  entries.forEach(e => {
    const safeRedaksi = e.redaksi.replace(/\s+/g, " ").trim();
    content += e.noUrut + "\t" + safeRedaksi + "\n";
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

// ====== DOWNLOAD WORD (.DOC) ======
document.getElementById("downloadDocBtn").addEventListener("click", function () {
  if (entries.length === 0) {
    alert("Belum ada redaksi yang dibuat.");
    return;
  }

  let html = `
  <html xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:w="urn:schemas-microsoft-com:office:word"
        xmlns="http://www.w3.org/TR/REC-html40">
  <head><meta charset="utf-8"><title>Redaksi Taksiran</title></head>
  <body>
    <h2 style="font-family:Calibri;">Daftar Redaksi Taksiran</h2>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; font-family:Calibri; font-size:11pt;">
      <thead>
        <tr>
          <th>No.</th>
          <th>Redaksi</th>
        </tr>
      </thead>
      <tbody>
  `;

  entries.forEach(e => {
    const htmlRedaksi = segmentsToHtml(e.segments); // mempertahankan italic
    html += `
        <tr>
          <td>${e.noUrut}</td>
          <td>${htmlRedaksi}</td>
        </tr>
    `;
  });

  html += `
      </tbody>
    </table>
  </body></html>`;

  const blob = new Blob([html], { type: "application/msword;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "redaksi_taksiran.doc";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

// ====== INIT: MUAT RIWAYAT DARI LOCALSTORAGE ======
loadFromStorage();
