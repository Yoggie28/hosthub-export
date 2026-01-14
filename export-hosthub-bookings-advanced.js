const axios = require("axios");
const ExcelJS = require("exceljs");
const fs = require("fs");
const readline = require("readline");

// ============== CONFIG ===================
const API_KEY = "ZGY4ODM5OTktN2Y3MC00NmVkLTg1M2MtMjI4NTJiNWVmZGRk"; // <-- Put your Hosthub API Key here
const BASE_URL = "https://app.hosthub.com/api/2019-03-01";
const DEFAULT_YEAR = 2025;
// ========================================

const api = axios.create({
  baseURL: BASE_URL,
  headers: {
    Authorization: API_KEY,
    "Content-Type": "application/json",
  },
});

// Convert cents to money
function money(value) {
  if (!value || value.cents == null) return 0;
  return value.cents / 100;
}

// Map source names to human-readable channels
function getChannel(source) {
  if (!source || !source.name) return "Offline";
  if (source.name.toLowerCase().includes("cityden2")) return "Booking";
  if (source.name.toLowerCase().includes("cityden")) return "Airbnb";
  return "Offline";
}

// Fetch all paginated data
async function fetchAll(url) {
  let results = [];
  let next = url;
  let page = 1;

  while (next) {
    const res = await api.get(next.replace(BASE_URL, ""));
    const dataLength = res.data.data.length;

    console.log(`  ➤ Fetched page ${page}, ${dataLength} events`);
    results.push(...res.data.data);

    if (dataLength === 0) break; // stop if no events
    next = res.data.navigation?.next || null;
    page++;
  }

  return results;
}

// Fetch Greek taxes for a booking
async function fetchGreekTaxes(calendarEventId) {
  try {
    const res = await api.get(
      `/calendar-events/${calendarEventId}/calendar-event-gr-taxes`
    );
    return res.data;
  } catch {
    return null;
  }
}

// Simple readline question wrapper
async function ask(question) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  return new Promise((resolve) => rl.question(question, (ans) => {
    rl.close();
    resolve(ans.trim());
  }));
}

async function main() {
  console.log("📥 Fetching rentals...");
  const rentals = (await api.get("/rentals")).data.data;

  rentals.forEach((r, i) => console.log(`${i + 1}. ${r.name}`));

  const rentalIndex = parseInt(await ask("Enter rental number to export: "), 10) - 1;
  const monthFilter = await ask("Enter month number (1-12) to filter or leave empty for all: ");
  const yearFilter = await ask(`Enter year to filter or leave empty for default ${DEFAULT_YEAR}: `);

  const rental = rentals[rentalIndex];
  const month = monthFilter ? parseInt(monthFilter, 10) : null;
  const year = yearFilter ? parseInt(yearFilter, 10) : DEFAULT_YEAR;

  console.log(`\n🏠 Exporting bookings for: ${rental.name}`);
  console.log("📥 Fetching calendar events...");

  const events = await fetchAll(`/rentals/${rental.id}/calendar-events`);

  let bookings = events.filter(e => 
    e.type === "Booking" &&
    e.is_visible &&
    new Date(e.date_from).getFullYear() === year &&
    (!month || new Date(e.date_from).getMonth() + 1 === month)
  );

  // Sort by Check-in ascending
  bookings.sort((a, b) => new Date(a.date_from) - new Date(b.date_from));

  if (bookings.length === 0) {
    console.log("  ↳ No bookings found for this month/year.");
    return;
  }

  if (!fs.existsSync("exports")) fs.mkdirSync("exports");

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Bookings");

  sheet.columns = [
    { header: "Reservation ID", width: 20 },
    { header: "Guest Name", width: 25 },
    { header: "Check-in", width: 12 },
    { header: "Check-out", width: 12 },
    { header: "Nights", width: 8 },
    { header: "Total Value", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "Cleaning Fee", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "Taxes", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "Guest Paid", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "Host Payout", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "VAT", width: 12, style: { numFmt: "#,##0.00" } },
    { header: "Climate Tax", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "Accommodation Tax", width: 18, style: { numFmt: "#,##0.00" } },
    { header: "AADE Value", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "Currency", width: 10 },
    { header: "Channel", width: 15 },
  ];

  // Totals accumulators
  let totalValue = 0;
  let totalCleaningFee = 0;
  let totalTaxes = 0;
  let totalGuestPaid = 0;
  let totalHostPayout = 0;

  for (const b of bookings) {
    const tax = await fetchGreekTaxes(b.id);

    const rowValues = [
      b.reservation_id || "",
      b.guest_name || "",
      b.date_from,
      b.date_to,
      b.nights,
      money(b.total_value), // <-- total value instead of booking value
      money(b.cleaning_fee),
      money(b.taxes),
      money(b.guest_paid),
      money(b.total_payout),
      money(tax?.vat),
      money(tax?.climate_tax),
      money(tax?.accommodation_tax),
      money(tax?.aade_value),
      b.booking_value?.currency || "",
      getChannel(b.source),
    ];

    totalValue += money(b.total_value);
    totalCleaningFee += money(b.cleaning_fee);
    totalTaxes += money(b.taxes);
    totalGuestPaid += money(b.guest_paid);
    totalHostPayout += money(b.total_payout);

    sheet.addRow(rowValues);
  }

  // Add totals row with formulas
  const lastRowNumber = sheet.rowCount;
  sheet.addRow([
    "TOTALS", "", "", "", "",
    { formula: `SUM(F2:F${lastRowNumber})`, style: { numFmt: "#,##0.00" } },
    { formula: `SUM(G2:G${lastRowNumber})`, style: { numFmt: "#,##0.00" } },
    { formula: `SUM(H2:H${lastRowNumber})`, style: { numFmt: "#,##0.00" } },
    { formula: `SUM(I2:I${lastRowNumber})`, style: { numFmt: "#,##0.00" } },
    { formula: `SUM(J2:J${lastRowNumber})`, style: { numFmt: "#,##0.00" } },
    { formula: `SUM(K2:K${lastRowNumber})`, style: { numFmt: "#,##0.00" } },
    { formula: `SUM(L2:L${lastRowNumber})`, style: { numFmt: "#,##0.00" } },
    { formula: `SUM(M2:M${lastRowNumber})`, style: { numFmt: "#,##0.00" } },
    { formula: `SUM(N2:N${lastRowNumber})`, style: { numFmt: "#,##0.00" } },
    "", ""
  ]);

  const safeName = rental.name.replace(/[\/\\:*?"<>|]/g, "");
  const filename = `exports/${safeName}-${year}.xlsx`;

  await workbook.xlsx.writeFile(filename);
  console.log(`  ✅ Exported: ${filename}`);
  console.log("🎉 Export completed!");
}

main().catch(err => console.error("❌ Error:", err.response?.data || err.message));
