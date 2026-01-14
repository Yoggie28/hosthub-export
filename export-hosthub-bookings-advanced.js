const axios = require("axios");
const ExcelJS = require("exceljs");
const fs = require("fs");
const readline = require("readline");
const { default: openFile } = require("open"); // Make sure: npm install open

// ============== CONFIG ===================
const API_KEY = "ZGY4ODM5OTktN2Y3MC00NmVkLTg1M2MtMjI4NTJiNWVmZGRk"; 
const BASE_URL = "https://app.hosthub.com/api/2019-03-01";
const DEFAULT_YEAR = 2025;
const CONFIG_FILE = "./config.json"; 
// ========================================

const api = axios.create({
  baseURL: BASE_URL,
  headers: {
    Authorization: API_KEY,
    "Content-Type": "application/json",
  },
});

// Load accommodation config
let accommodationsConfig = {};
if (fs.existsSync(CONFIG_FILE)) {
  accommodationsConfig = JSON.parse(fs.readFileSync(CONFIG_FILE, "utf8"));
}

// Convert cents to money
function money(value) {
  if (!value || value.cents == null) return 0;
  return value.cents / 100;
}

// Format YYYY-MM-DD → DD/MM/YYYY
function formatDate(dateStr) {
  const d = new Date(dateStr);
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
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
    if (dataLength === 0) break;
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
  return new Promise((resolve) =>
    rl.question(question, (ans) => {
      rl.close();
      resolve(ans.trim());
    })
  );
}

async function main() {
  console.log("📥 Fetching rentals...");
  const rentals = (await api.get("/rentals")).data.data;

  rentals.forEach((r, i) => console.log(`${i + 1}. ${r.name}`));

  const rentalIndex = parseInt(
    await ask("Enter rental number to export: "),
    10
  ) - 1;
  const monthFilter = await ask(
    "Enter month number (1-12) to filter or leave empty for all: "
  );
  const yearFilter = await ask(
    `Enter year to filter or leave empty for default ${DEFAULT_YEAR}: `
  );

  const rental = rentals[rentalIndex];
  const month = monthFilter ? parseInt(monthFilter, 10) : null;
  const year = yearFilter ? parseInt(yearFilter, 10) : DEFAULT_YEAR;

  console.log(`\n🏠 Exporting bookings for: ${rental.name}`);
  console.log("📥 Fetching calendar events...");

  let events = await fetchAll(`/rentals/${rental.id}/calendar-events`);

  let bookings = events.filter(
    (e) =>
      e.type === "Booking" &&
      e.is_visible &&
      new Date(e.date_from).getFullYear() === year &&
      (!month || new Date(e.date_from).getMonth() + 1 === month)
  );

  bookings.sort((a, b) => new Date(a.date_from) - new Date(b.date_from));

  if (bookings.length === 0) {
    console.log("  ↳ No bookings found for this month/year.");
    return;
  }

  if (!fs.existsSync("exports")) fs.mkdirSync("exports");

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Bookings");

  // Columns (no per-row Grand Total)
  sheet.columns = [
    { header: "Channel", width: 15 },
    { header: "Check-in → Check-out", width: 25 },
    { header: "Guest Name", width: 25 },
    { header: "Total Value", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "Climate Tax", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "Cleaning Fee", width: 15, style: { numFmt: "#,##0.00" } },
    { header: "Commission Amount", width: 18, style: { numFmt: "#,##0.00" } },
    { header: "Commission", width: 15, style: { numFmt: "#,##0.00" } }
  ];

  let totalValue = 0;
  let totalClimateTax = 0;
  let totalCleaningFee = 0;
  let totalCommissionAmount = 0;
  let totalCommission = 0;

  for (const b of bookings) {
    const tax = await fetchGreekTaxes(b.id);
    const config = accommodationsConfig[rental.name] || { cleaningFee: 0, commissionPercent: 0 };

    const channel = getChannel(b.source);
    const cleaningFee = channel === "Offline" ? config.cleaningFee : money(b.cleaning_fee);
    const climateTax = channel === "Offline" ? 0 : money(tax?.climate_tax);
    const totalVal = money(b.total_value);

    const commissionAmount = totalVal - climateTax - cleaningFee;
    const commission = commissionAmount * (config.commissionPercent / 100);

    totalValue += totalVal;
    totalClimateTax += climateTax;
    totalCleaningFee += cleaningFee;
    totalCommissionAmount += commissionAmount;
    totalCommission += commission;

    sheet.addRow([
      channel,
      `${formatDate(b.date_from)} → ${formatDate(b.date_to)}`,
      b.guest_name || "",
      totalVal,
      climateTax,
      cleaningFee,
      commissionAmount,
      commission
    ]);
  }

  // Totals row
  sheet.addRow([
    "TOTALS",
    "",
    "",
    totalValue,
    totalClimateTax,
    totalCleaningFee,
    totalCommissionAmount,
    totalCommission
  ]);

  // GRAND TOTAL row (Cleaning Fee + Commission)
  sheet.addRow([
    "GRAND TOTAL",
    "",
    "",
    "",
    "",
    totalCleaningFee,
    "",
    totalCommission,
    totalCleaningFee + totalCommission
  ]);

  const safeName = rental.name.replace(/[\/\\:*?"<>|]/g, "");
  const filename = `exports/${safeName}-${year}.xlsx`;

  await workbook.xlsx.writeFile(filename);
  console.log(`  ✅ Exported: ${filename}`);

  // Open Excel file
  await openFile(filename);

  console.log("🎉 Export completed!");
}

main().catch((err) =>
  console.error("❌ Error:", err.response?.data || err.message)
);
