// netlify/functions/redaksi.mts
import { neon } from "@netlify/neon";
import type { Context, Config } from "@netlify/functions";

// koneksi ke Netlify DB (Neon)
const sql = neon();

// pastikan tabel ada (aman kalau dipanggil berulang)
async function ensureTable() {
  await sql(`
    CREATE TABLE IF NOT EXISTS redaksi (
      id SERIAL PRIMARY KEY,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      no_urut TEXT,
      barang TEXT,
      karat INTEGER,
      berat NUMERIC(10,2),
      harga BIGINT,
      redaksi_text TEXT NOT NULL,
      raw_data JSONB
    )
  `);
}

// handler utama
export default async (req: Request, context: Context) => {
  await ensureTable();

  // GET  → ambil daftar redaksi (untuk Riwayat & Dashboard)
  if (req.method === "GET") {
    const rows = await sql<{
      id: number;
      created_at: string;
      no_urut: string | null;
      barang: string | null;
      karat: number | null;
      berat: string | null;
      harga: string | null;
      redaksi_text: string;
      raw_data: any;
    }[]>`
      SELECT *
      FROM redaksi
      ORDER BY created_at DESC
      LIMIT 200
    `;

    return new Response(JSON.stringify(rows), {
      headers: { "Content-Type": "application/json" },
    });
  }

  // POST → simpan redaksi baru
  if (req.method === "POST") {
    let body: any;
    try {
      body = await req.json();
    } catch {
      return new Response(JSON.stringify({ error: "Body harus JSON" }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }

    const {
      no_urut,
      barang,
      karat,
      berat,
      harga,
      redaksi_text,
      raw_data,
    } = body || {};

    if (!redaksi_text) {
      return new Response(JSON.stringify({ error: "redaksi_text wajib diisi" }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }

    await sql`
      INSERT INTO redaksi
        (no_urut, barang, karat, berat, harga, redaksi_text, raw_data)
      VALUES
        (
          ${no_urut || null},
          ${barang || null},
          ${karat ?? null},
          ${berat ?? null},
          ${harga ?? null},
          ${redaksi_text},
          ${raw_data ? JSON.stringify(raw_data) : null}
        )
    `;

    return new Response(JSON.stringify({ success: true }), {
      headers: { "Content-Type": "application/json" },
    });
  }

  // metode lain tidak diizinkan
  return new Response("Method Not Allowed", { status: 405 });
};

// supaya URL-nya rapi → /api/redaksi
export const config: Config = {
  path: "/api/redaksi",
};
