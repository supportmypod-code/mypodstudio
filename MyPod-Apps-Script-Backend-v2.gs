/**
 * MyPod booking backend v2 (Google Apps Script + Google Sheets)
 *
 * What is new compared with v1
 *  - getSlots: one call returns every time slot for a date with availability (the site shows slot buttons).
 *  - Prices and coupons are calculated here. The browser can no longer set its own price.
 *  - Coupons sheet: add, pause or expire offers without touching code.
 *  - lookupBooking / cancelBooking: customers manage a booking with booking ID + phone.
 *  - Anti-spam: hidden honeypot field plus a per-phone hourly limit.
 *  - Owner alert email on every new booking and cancellation (set OWNER_EMAIL).
 *  - Fix: date, time and phone columns are stored as plain text. In v1 Google Sheets turned
 *    "8:30 AM" into a time value, so the double-booking check could stop matching.
 *
 * v2.2: A coupon is REQUIRED for every booking (doorstep and drive-through). Coupons come from the 1200-coupon sheet
 *       (MYPOD_300_Flyers_1200_Coupons_Backend.xlsx, tab "All Coupons"). Each coupon works once: the booking marks it Redeemed (with the customer,
 *       booking ID and time) and a cancelled booking gives the coupon back.
 *
 * v2.3: Customer login and reminders.
 *  - Customers sign in on the website with their booking email: requestOtp sends a 6-digit code, verifyOtp returns a 30-day session token,
 *    myBookings lists their bookings, cancelMy cancels one. No passwords are stored. Codes expire after 10 minutes.
 *  - sendReminders() emails each customer about 24 hours and 1 hour before their slot. RUN installReminders() ONCE from the editor,
 *    approve the new permission (it needs to create a timer), and reminders then run by themselves every 15 minutes.
 *  - After pasting this version: Deploy > Manage deployments > New version.
 *
 * SETUP: same as v1.
 *  1. Bind this script to the "MyPod Bookings" sheet (Extensions > Apps Script).
 *  2. Run setupSheet() once and approve permissions.
 *  3. Deploy > New deployment > Web app. Execute as: Me. Access: Anyone.
 *  3b. v2.1: bookings now save the pincode/area and a Google Maps link automatically (the Bookings sheet has one extra column,
 *      "Map Link"). If you already ran an earlier setup, rename the old Bookings sheet first. After pasting this version, run any
 *      function once and approve the new Google Maps permission, then Deploy > Manage deployments > New version.
 *  4. If you already have a v1 "Bookings" sheet, rename it (for example "Bookings v1") before running setupSheet(),
 *     because v2 uses more columns.
 */

const SHEET_NAME = "Bookings";
const COUPON_SHEET = "All Coupons";      // tab imported from MYPOD_300_Flyers_1200_Coupons_Backend.xlsx
const TIME_ZONE = "Asia/Kolkata";
const OWNER_EMAIL = "supportmypod@gmail.com";                 // e.g. "owner@mypod.in". Leave empty to skip email alerts.
const MAX_BOOKINGS_PER_PHONE_PER_HOUR = 4;

// Drive-through location shown to customers and saved on the booking.
const BRAND_NAME = "MyPod Car Care";
const SUPPORT_WHATSAPP = "919886950617";   // digits only, used in the customer email
const SITE_URL = "https://supportmypod-code.github.io/mypodstudio/";   // Adds a "manage or cancel" link to the customer email.
const SEND_CUSTOMER_EMAIL = true;           // Gmail allows about 100 emails a day, Google Workspace about 1,500.
const STUDIO_ADDRESS = "MYPOD DETAILING STUDIO, Action Area I, Newtown, Kolkata";
// Doorstep service area [lat, lng], traced from the MyPod coverage map. Edit the points to widen or narrow it.
const SERVICE_AREA = [[22.61321,88.48137],[22.60794,88.48192],[22.60405,88.48953],[22.59376,88.49483],[22.589,88.50352],[22.58661,88.50937],[22.58862,88.51575],[22.58611,88.52608],[22.57934,88.52853],[22.57219,88.5341],[22.56416,88.53791],[22.55801,88.54253],[22.55199,88.53315],[22.53844,88.52133],[22.53744,88.51494],[22.53944,88.51018],[22.54007,88.50543],[22.53932,88.49877],[22.5407,88.4951],[22.53844,88.49102],[22.54208,88.48871],[22.55362,88.47118],[22.55914,88.46466],[22.57683,88.44672],[22.57909,88.44536],[22.58523,88.44495],[22.58636,88.44264],[22.5959,88.44291],[22.60041,88.44699],[22.60167,88.44971],[22.60606,88.45786],[22.60919,88.46642],[22.60606,88.47349],[22.60531,88.47757]];
// Extra doorstep area: everything within radius (metres) of this point. 22°38'50.5"N 88°22'51.2"E
const EXTRA_AREAS = [{ lat: 22.647361, lng: 88.380889, radius: 1000 }];
const NEWTOWN_RE = /(newtown|new town|rajarhat|action area|700156|700157|700135|700160)/i;

// Bookings allowed at the same date + time + service mode.
const SLOT_CAPACITY = { "Doorstep": 1, "Drive-through": 1 };

const SLOTS = ["8:30 AM","9:15 AM","10:00 AM","10:45 AM","11:30 AM","12:15 PM",
               "1:00 PM","1:45 PM","2:30 PM","3:15 PM","4:00 PM","4:45 PM","5:30 PM","6:00 PM"];

// Source of truth for prices. null = quote confirmed by MyPod. Keep in sync with the website.
const CATALOG = {
  "Car":  { "Promotional Car Wash":299, "Ceramic Protection":null, "Paint Protection Film (PPF)":null },
  "Bike": { "Promotional Bike Wash":99 }
};

const HEADERS = ["Created At","Booking ID","Status","Customer Name","Phone","Service Mode","Area / Pincode",
  "Address","Landmark","Vehicle Type","Service","Coupon Code","List Price (INR)","Discount (INR)",
  "Total (INR)","Date","Time","Cancelled At","Map Link","Email"];
const TEXT_COLUMNS = ["Phone","Date","Time","Booking ID","Coupon Code"];

// Website service name -> service name used in the Coupons sheet (column "service").
const COUPON_SERVICE = {
  "Promotional Car Wash": "Car Washing",
  "Ceramic Protection": "Ceramic Coating",
  "Paint Protection Film (PPF)": "Paint Protection Film (PPF)",
  "Promotional Bike Wash": "Bike Washing"
};
const COUPON_HEADERS = ["flyer_id","flyer_number","coupon_id","coupon_code","brand","service","offer_type","offer_value","valid_until","status",
  "usage_limit","used_count","customer_name","customer_phone","booking_id","issued_at","redeemed_at","notes"];
const MAX_BAD_COUPON_TRIES = 40;           // wrong-code attempts allowed per 15 minutes (all visitors together), stops code guessing

/* ---------- Web app entry points ---------- */

function doGet() {
  return json_({ ok: true, service: "MyPod booking API v2", status: "ready" });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (body.website) return json_({ ok: true });            // honeypot field filled in by a bot
    switch (String(body.action || "")) {
      case "getSlots":        return json_(getSlots_(body));
      case "validateCoupon":  return json_(validateCoupon_(body));
      case "createBooking":   return json_(createBooking_(body));
      case "lookupBooking":   return json_(lookupBooking_(body));
      case "cancelBooking":   return json_(cancelBooking_(body));
      case "requestOtp":      return json_(requestOtp_(body));
      case "verifyOtp":       return json_(verifyOtp_(body));
      case "myBookings":      return json_(myBookings_(body));
      case "cancelMy":        return json_(cancelMy_(body));
      default:                return json_({ ok: false, message: "Unsupported action." });
    }
  } catch (err) {
    return json_({ ok: false, message: err && err.message ? err.message : "Unexpected booking error." });
  }
}

/* ---------- Actions ---------- */

function getSlots_(b) {
  requireFields_(b, ["serviceMode", "date"]);
  checkMode_(b.serviceMode); checkDate_(b.date);
  const taken = {};
  readBookings_().forEach(r => {
    if (isActive_(r.status) && r.mode === b.serviceMode && r.date === b.date) taken[r.time] = (taken[r.time] || 0) + 1;
  });
  const cap = SLOT_CAPACITY[b.serviceMode];
  return { ok: true, slots: SLOTS.map(t => ({ time: t, available: (taken[t] || 0) < cap })) };
}

function validateCoupon_(b) {
  requireFields_(b, ["couponCode", "vehicleType", "serviceName"]);
  const price = priceFor_(b.vehicleType, b.serviceName);
  const c = checkCoupon_(b.couponCode, b.serviceName);
  const discount = discountFor_(c, price);
  return { ok: true, code: c.code, discount, message: c.label + (price === null ? " (applied to the final quote)" : "") };
}

function createBooking_(b) {
  requireFields_(b, ["serviceMode","vehicleType","serviceName","date","time","customerName","customerPhone","customerEmail"]);
  if (!String(b.couponCode || "").trim()) throw new Error("A valid coupon code is required to book.");   // doorstep and drive-through
  checkMode_(b.serviceMode); checkDate_(b.date);
  if (["Car","Bike"].indexOf(b.vehicleType) < 0) throw new Error("Invalid vehicle type.");
  if (SLOTS.indexOf(String(b.time)) < 0) throw new Error("Invalid time slot.");
  if (b.serviceMode === "Doorstep" && !String(b.address || "").trim()) throw new Error("Doorstep address is required.");
  const where = b.serviceMode === "Doorstep" ? resolveArea_(b) : { area: "Drive-through", map: "" };   // pincode and area are fetched here, not typed by the customer
  const phone = normalizePhone_(b.customerPhone);
  const email = normalizeEmail_(b.customerEmail);
  if (String(b.customerName).length > 120) throw new Error("Name is too long.");

  const price = priceFor_(b.vehicleType, b.serviceName);          // server-side price
  checkCoupon_(b.couponCode, b.serviceName);                       // fail early with a clear message; checked again inside the lock

  rateLimit_(phone);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const cap = SLOT_CAPACITY[b.serviceMode];
    const used = readBookings_().filter(r => isActive_(r.status) && r.mode === b.serviceMode && r.date === b.date && r.time === b.time).length;
    if (used >= cap) throw new Error("Sorry, this slot was just booked. Please choose another time.");

    const coupon = checkCoupon_(b.couponCode, b.serviceName);      // re-check under the lock so one coupon cannot be used twice
    const discount = discountFor_(coupon, price), couponCode = coupon.code;

    const bookingId = "MP-" + Utilities.formatDate(new Date(), TIME_ZONE, "yyyyMMdd") + "-" +
      Utilities.getUuid().replace(/-/g, "").slice(0, 6).toUpperCase();
    const total = price === null ? "" : Math.max(0, price - discount);
    const createdAt = Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
    const row = {
      "Created At": createdAt, "Booking ID": bookingId, "Status": "Confirmed",
      "Customer Name": clean_(b.customerName), "Phone": phone, "Service Mode": b.serviceMode,
      "Area / Pincode": clean_(where.area),
      "Address": b.serviceMode === "Doorstep" ? clean_(b.address) : STUDIO_ADDRESS,
      "Landmark": clean_(b.landmark), "Vehicle Type": b.vehicleType, "Service": b.serviceName,
      "Coupon Code": couponCode, "List Price (INR)": price === null ? "" : price, "Discount (INR)": discount,
      "Total (INR)": total, "Date": b.date, "Time": b.time, "Cancelled At": "", "Map Link": where.map, "Email": email
    };
    appendRow_(row);
    redeemCoupon_(coupon, bookingId, clean_(b.customerName), phone);
    notifyOwner_("New booking " + bookingId, row);
    sendCustomerEmail_("confirmed", row);
    return { ok: true, booking: publicBooking_(row) };
  } finally {
    lock.releaseLock();
  }
}

function lookupBooking_(b) {
  requireFields_(b, ["bookingId", "customerPhone"]);
  const hit = findBooking_(b.bookingId, b.customerPhone);
  return { ok: true, booking: publicBooking_(hit.obj) };
}

function cancelBooking_(b) {
  requireFields_(b, ["bookingId", "customerPhone"]);
  return cancelHit_(() => findBooking_(b.bookingId, b.customerPhone));
}

// Signed-in customer cancels one of their own bookings (matched by the email in their session token).
function cancelMy_(b) {
  const email = readToken_(b.token);
  requireFields_(b, ["bookingId"]);
  return cancelHit_(() => {
    const id = String(b.bookingId).trim().toUpperCase();
    const hit = readBookings_().filter(r => String(r.obj["Booking ID"]).toUpperCase() === id && emailOf_(r.obj) === email)[0];
    if (!hit) throw new Error("No booking found.");
    return hit;
  });
}

function cancelHit_(find) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const hit = find();
    if (!isActive_(hit.obj["Status"])) throw new Error("This booking is already " + hit.obj["Status"].toLowerCase() + ".");
    const start = new Date(hit.obj["Date"] + "T00:00:00+05:30");
    if (start < new Date(Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd") + "T00:00:00+05:30"))
      throw new Error("Past bookings cannot be cancelled.");
    const sheet = getSheet_();
    sheet.getRange(hit.rowNumber, HEADERS.indexOf("Status") + 1).setValue("Cancelled");
    sheet.getRange(hit.rowNumber, HEADERS.indexOf("Cancelled At") + 1)
      .setValue(Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd HH:mm:ss"));
    hit.obj["Status"] = "Cancelled";
    releaseCoupon_(hit.obj["Coupon Code"], hit.obj["Booking ID"]);   // the customer did not get the service, so the coupon is usable again
    notifyOwner_("Cancelled " + hit.obj["Booking ID"], hit.obj);
    sendCustomerEmail_("cancelled", hit.obj);
    return { ok: true, booking: publicBooking_(hit.obj) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Customer login (email one-time code) ---------- */

const OTP_MINUTES = 10;                    // how long a code works
const SESSION_DAYS = 30;                   // how long a customer stays signed in
const MAX_OTP_REQUESTS_PER_15_MIN = 3;     // codes per email address
const MAX_OTP_GUESSES = 5;                 // wrong codes before a new one is needed

function emailOf_(o) { return String(o["Email"] || "").trim().toLowerCase(); }

function secret_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty("SESSION_SECRET");
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); props.setProperty("SESSION_SECRET", k); }
  return k;
}

function sign_(text) {
  return Utilities.computeHmacSha256Signature(String(text), secret_()).map(x => ("0" + (x & 255).toString(16)).slice(-2)).join("");
}

function makeToken_(email) {
  const body = Utilities.base64EncodeWebSafe(email + "|" + (Date.now() + SESSION_DAYS * 864e5));
  return body + "." + sign_(body);
}

function readToken_(token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 2 || sign_(parts[0]) !== parts[1]) throw new Error("Please sign in again.");
  const bits = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString().split("|");
  if (Number(bits[1]) < Date.now()) throw new Error("Your session expired. Please sign in again.");
  return bits[0];
}

function requestOtp_(b) {
  requireFields_(b, ["email"]);
  const email = normalizeEmail_(b.email), cache = CacheService.getScriptCache();
  const rl = "otprl_" + email, n = Number(cache.get(rl) || 0);
  if (n >= MAX_OTP_REQUESTS_PER_15_MIN) throw new Error("Too many codes requested. Please wait a few minutes and try again.");
  cache.put(rl, String(n + 1), 900);
  // The answer is the same whether or not the email has a booking, so nobody can use this to find out who our customers are.
  if (readBookings_().some(r => emailOf_(r.obj) === email)) {
    const code = String(100000 + parseInt(Utilities.getUuid().replace(/-/g, "").slice(0, 8), 16) % 900000);
    cache.put("otp_" + email, sign_(email + ":" + code), OTP_MINUTES * 60);
    cache.remove("otpg_" + email);
    try {
      MailApp.sendEmail(email, "Your MyPod sign-in code: " + code,
        "Your MyPod sign-in code is " + code + ". It works for " + OTP_MINUTES + " minutes. If you did not ask for it, you can ignore this email.",
        { name: BRAND_NAME, replyTo: OWNER_EMAIL || undefined,
          htmlBody: '<div style="background:#050505;padding:24px 12px;font-family:Arial,Helvetica,sans-serif"><div style="max-width:420px;margin:0 auto;background:#0e0e0e;border:1px solid #2a2a2a;border-radius:16px;padding:26px;text-align:center">' +
            '<div style="font-size:24px;font-weight:900;color:#fff">MY<span style="color:#ffd51b">POD</span></div>' +
            '<p style="color:#d8d8d8;font-size:14px;margin:18px 0 8px">Your sign-in code</p>' +
            '<div style="font-size:36px;font-weight:800;letter-spacing:8px;color:#ffd51b">' + code + '</div>' +
            '<p style="color:#8a8a8a;font-size:12px;margin:18px 0 0">Works for ' + OTP_MINUTES + ' minutes. If you did not ask for it, ignore this email.</p></div></div>' });
    } catch (e) { /* the customer can ask for another code */ }
  }
  return { ok: true, message: "If that email has a booking, a 6-digit code is on its way. It works for " + OTP_MINUTES + " minutes." };
}

function verifyOtp_(b) {
  requireFields_(b, ["email", "code"]);
  const email = normalizeEmail_(b.email), cache = CacheService.getScriptCache();
  const gk = "otpg_" + email, guesses = Number(cache.get(gk) || 0);
  if (guesses >= MAX_OTP_GUESSES) { cache.remove("otp_" + email); throw new Error("Too many wrong codes. Please ask for a new code."); }
  const want = cache.get("otp_" + email);
  if (!want) throw new Error("That code has expired or was not sent. Please ask for a new code.");
  if (sign_(email + ":" + String(b.code).trim()) !== want) { cache.put(gk, String(guesses + 1), 900); throw new Error("That code is not correct."); }
  cache.remove("otp_" + email); cache.remove(gk);
  return { ok: true, token: makeToken_(email), email: email };
}

function myBookings_(b) {
  const email = readToken_(b.token);
  const bookings = readBookings_().filter(r => emailOf_(r.obj) === email).map(r => {
    const x = publicBooking_(r.obj);
    x.mapLink = String(r.obj["Map Link"] || "");
    return x;
  });
  return { ok: true, email: email, bookings: bookings };
}

/* ---------- Reminders (automatic emails before the slot) ---------- */

// Run this ONCE from the editor. It sets up a timer that calls sendReminders every 15 minutes.
function installReminders() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "sendReminders").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("sendReminders").timeBased().everyMinutes(15).create();
  return "Done. Reminder emails now go out automatically (checked every 15 minutes).";
}

function startMs_(o) {
  const d = normalizeDate_(o["Date"]), t = /(\d+):(\d+) (AM|PM)/.exec(normalizeTime_(o["Time"]));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !t) return NaN;
  let h = Number(t[1]) % 12; if (t[3] === "PM") h += 12;
  return new Date(d + "T" + ("0" + h).slice(-2) + ":" + t[2] + ":00+05:30").getTime();
}

function sendReminders() {
  const props = PropertiesService.getScriptProperties(), now = Date.now(), H = 3600000;
  readBookings_().forEach(r => {
    if (!isActive_(r.status) || !r.obj["Email"]) return;
    const start = startMs_(r.obj), left = start - now;
    if (!(left > 0)) return;
    [["24h", left <= 24 * H && left > 3 * H], ["1h", left <= 75 * 60000]].forEach(k => {
      const key = "rem_" + r.obj["Booking ID"] + "_" + k[0];
      if (!k[1] || props.getProperty(key)) return;
      props.setProperty(key, String(start));          // mark first, so a customer never gets the same reminder twice
      sendReminderEmail_(k[0], r.obj);
    });
  });
  const all = props.getProperties();                  // tidy up old markers
  Object.keys(all).forEach(k => { if (k.indexOf("rem_") === 0 && Number(all[k]) < now - 2 * 864e5) props.deleteProperty(k); });
}

function sendReminderEmail_(kind, o) {
  try {
    const first = String(o["Customer Name"]).trim().split(/\s+/)[0], id = o["Booking ID"];
    const when = prettyDate_(normalizeDate_(o["Date"])) + " at " + normalizeTime_(o["Time"]);
    const head = kind === "1h" ? "Your slot starts in about an hour" : "Your MyPod slot is coming up";
    const doorstep = o["Service Mode"] === "Doorstep";
    const wa = "https://wa.me/" + SUPPORT_WHATSAPP + "?text=" + encodeURIComponent("Hi MyPod, about my booking " + id);
    const manage = SITE_URL ? '<p style="margin:14px 0 0;font-size:13px"><a href="' + esc_(SITE_URL) + '" style="color:#ffd51b">Sign in to see or cancel your booking</a></p>' : "";
    const html = '<div style="background:#050505;padding:24px 12px;font-family:Arial,Helvetica,sans-serif"><div style="max-width:520px;margin:0 auto;background:#0e0e0e;border:1px solid #2a2a2a;border-radius:16px;padding:26px">' +
      '<div style="font-size:24px;font-weight:900;color:#fff">MY<span style="color:#ffd51b">POD</span></div>' +
      '<h1 style="margin:18px 0 8px;font-size:22px;color:#ffd51b">' + head + '</h1>' +
      '<p style="margin:0 0 14px;color:#d8d8d8;font-size:14px;line-height:1.6">Hi ' + esc_(first) + ', a quick reminder about your booking.</p>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #2a2a2a;border-bottom:1px solid #2a2a2a;font-size:14px;color:#f4f4f4">' +
      '<tr><td style="padding:8px 0;color:#8a8a8a;width:32%">When</td><td style="font-weight:600">' + esc_(when) + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#8a8a8a">Service</td><td style="font-weight:600">' + esc_(o["Service"]) + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#8a8a8a">' + (doorstep ? "Address" : "Drive in at") + '</td><td style="font-weight:600">' + esc_(o["Address"]) + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#8a8a8a">Booking ID</td><td style="font-weight:600">' + esc_(id) + '</td></tr></table>' +
      '<p style="margin:16px 0 0;color:#cfcfcf;font-size:13px;line-height:1.6">' + (doorstep ? "Please keep the car accessible and your phone nearby." : "Please arrive a few minutes early.") + '</p>' +
      manage + '<p style="margin:16px 0 0"><a href="' + wa + '" style="display:inline-block;padding:11px 18px;background:#ffd51b;border-radius:24px;color:#111;text-decoration:none;font-size:13px;font-weight:800">Chat on WhatsApp</a></p></div></div>';
    MailApp.sendEmail(o["Email"], (kind === "1h" ? "Starting soon · " : "Reminder · ") + o["Service"] + " · " + when,
      head + "\n" + when + "\n" + o["Service"] + "\nBooking ID: " + id + "\nWhatsApp: " + wa,
      { name: BRAND_NAME, htmlBody: html, replyTo: OWNER_EMAIL || undefined });
  } catch (e) { /* ignore: a failed reminder never breaks anything */ }
}

/* ---------- Coupon tools (run from the Apps Script editor) ---------- */

const TEST_COUPON_CODE = "MPTEST-4829";   // single-use test coupon: 50% off any service, valid 7 days

// DELETES EVERY ROW in the "All Coupons" tab (the header row stays). Bookings already made are not touched.
function deleteAllCoupons() {
  const sh = loadCoupons_().sheet, last = sh.getLastRow();
  if (last > 2) sh.deleteRows(3, last - 2);                                       // Sheets will not delete every non-frozen row, so keep row 2 and empty it
  if (last > 1) sh.getRange(2, 1, 1, sh.getLastColumn()).clearContent();
  return "Deleted " + Math.max(0, last - 1) + " coupons. Customers cannot book until you add new coupons.";
}

// Adds ONE test coupon (TEST_COUPON_CODE). It works once: the first booking marks it Redeemed; cancelling that booking gives it back.
function createTestCoupon() {
  const d = loadCoupons_(), col = d.col, sh = d.sheet;
  for (let i = 1; i < d.rows.length; i++) if (String(d.rows[i][col.coupon_code]).trim().toUpperCase() === TEST_COUPON_CODE) return TEST_COUPON_CODE + " already exists.";
  const row = new Array(sh.getLastColumn()).fill("");
  const set = (h, v) => { row[col[h]] = v; };
  const expires = Utilities.formatDate(new Date(Date.now() + 7 * 864e5), TIME_ZONE, "yyyy-MM-dd");
  set("flyer_id", "TEST"); set("flyer_number", 0); set("coupon_id", "TEST-1"); set("coupon_code", TEST_COUPON_CODE); set("brand", "MyPod");
  set("service", "");                       // empty = valid for every service
  set("offer_type", "percent"); set("offer_value", 50); set("valid_until", expires); set("status", "Active");
  set("usage_limit", 1); set("used_count", 0); set("issued_at", Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd HH:mm:ss")); set("notes", "Test coupon");
  sh.appendRow(row);
  return "Test coupon " + TEST_COUPON_CODE + " created: 50% off, one use, valid until " + expires + ".";
}

// One click: delete every coupon, then add the test coupon.
function resetCouponsToTest() { return deleteAllCoupons() + " " + createTestCoupon(); }

/* ---------- Setup ---------- */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Open Apps Script from the MyPod Bookings Google Sheet.");

  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else if (String(sheet.getRange(1, 2).getValue()) === "Booking ID" && sheet.getLastColumn() === HEADERS.length - 1 &&
             String(sheet.getRange(1, HEADERS.length - 1).getValue()) === "Map Link") {
    // Existing sheet from before the Email column: add it, keep all bookings.
    sheet.getRange(1, HEADERS.length).setValue("Email").setFontWeight("bold");
  } else if (String(sheet.getRange(1, 2).getValue()) !== "Booking ID" || sheet.getLastColumn() !== HEADERS.length) {
    throw new Error("The existing Bookings sheet uses the v1 layout. Rename it, then run setupSheet() again.");
  }
  TEXT_COLUMNS.forEach(h => sheet.getRange(1, HEADERS.indexOf(h) + 1, sheet.getMaxRows(), 1).setNumberFormat("@"));
  sheet.autoResizeColumns(1, HEADERS.length);

  let cs = ss.getSheetByName(COUPON_SHEET);
  if (!cs) {
    cs = ss.insertSheet(COUPON_SHEET);
    cs.getRange(1, 1, 1, COUPON_HEADERS.length).setValues([COUPON_HEADERS]).setFontWeight("bold");
    cs.setFrozenRows(1);
    return "Bookings sheet is ready. The Coupons sheet is empty: import MYPOD_300_Flyers_1200_Coupons_Backend.xlsx (see README).";
  }
  if (cs.getRange(1, 1, 1, cs.getLastColumn()).getValues()[0].indexOf("coupon_code") < 0) {
    throw new Error("The \"" + COUPON_SHEET + "\" sheet has the wrong layout. Rename it, import MYPOD_300_Flyers_1200_Coupons_Backend.xlsx, then run setupSheet() again.");
  }
  return "Bookings and Coupons sheets are ready.";
}

/* ---------- Helpers ---------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("No active spreadsheet. Bind this script to the booking spreadsheet.");
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Run setupSheet() once to create the Bookings sheet.");
  return sheet;
}

// Reads all bookings as normalized objects (rowNumber is the 1-based sheet row).
function readBookings_() {
  const values = getSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    HEADERS.forEach((h, c) => { obj[h] = values[i][c]; });
    obj["Date"] = normalizeDate_(obj["Date"]);
    obj["Time"] = normalizeTime_(obj["Time"]);
    obj["Phone"] = String(obj["Phone"] || "");
    out.push({ rowNumber: i + 1, obj: obj, status: String(obj["Status"] || ""), mode: String(obj["Service Mode"] || ""),
               date: obj["Date"], time: obj["Time"] });
  }
  return out;
}

function findBooking_(bookingId, phoneInput) {
  const phone = normalizePhone_(phoneInput);
  const id = String(bookingId).trim().toUpperCase();
  const hit = readBookings_().filter(r => String(r.obj["Booking ID"]).toUpperCase() === id && r.obj["Phone"].slice(-10) === phone)[0];
  if (!hit) throw new Error("No booking found for that ID and phone number.");
  return hit;
}

function appendRow_(rowObj) {
  const sheet = getSheet_();
  const rowNumber = sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([HEADERS.map(h => rowObj[h] === undefined ? "" : rowObj[h])]);
}

function publicBooking_(o) {
  return {
    bookingId: o["Booking ID"], status: o["Status"], customerName: o["Customer Name"], serviceMode: o["Service Mode"],
    address: o["Address"], vehicleType: o["Vehicle Type"], serviceName: o["Service"], couponCode: o["Coupon Code"],
    listPrice: o["List Price (INR)"] === "" ? null : Number(o["List Price (INR)"]),
    discount: Number(o["Discount (INR)"] || 0),
    price: o["Total (INR)"] === "" ? null : Number(o["Total (INR)"]),
    date: normalizeDate_(o["Date"]), time: normalizeTime_(o["Time"])
  };
}

function priceFor_(vehicle, service) {
  const group = CATALOG[String(vehicle)];
  if (!group || !(service in group)) throw new Error("Unknown service for this vehicle.");
  return group[service];
}

// Reads the Coupons sheet and finds one code. Columns are found by header name, so their order does not matter.
function loadCoupons_() {
  const cs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COUPON_SHEET);
  if (!cs) throw new Error("Coupons are not set up yet.");
  const rows = cs.getDataRange().getValues();
  const col = {};
  (rows[0] || []).forEach((h, i) => { col[String(h).trim()] = i; });
  COUPON_HEADERS.forEach(h => { if (col[h] === undefined) throw new Error("The \"" + COUPON_SHEET + "\" sheet is missing the column: " + h); });
  return { sheet: cs, rows: rows, col: col };
}

// Validates a coupon for a service WITHOUT using it. Throws a clear message when it cannot be used.
function checkCoupon_(codeInput, serviceName) {
  const code = normalizeCouponCode_(codeInput);
  if (!code) throw new Error("Enter your coupon code.");
  const data = loadCoupons_(), col = data.col, rows = data.rows;
  const today = Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd");
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[col.coupon_code]).trim().toUpperCase() !== code) continue;
    const status = String(r[col.status]).trim().toLowerCase();
    const limit = Number(r[col.usage_limit] || 1), used = Number(r[col.used_count] || 0);
    if (status === "redeemed" || used >= limit) throw new Error("This coupon has already been used.");
    if (status !== "active") throw new Error("This coupon is not active.");
    const expires = normalizeDate_(r[col.valid_until]);
    if (expires && expires < today) throw new Error("This coupon has expired.");
    const scope = String(r[col.service] || "").trim();
    if (scope && scope !== (COUPON_SERVICE[serviceName] || serviceName))
      throw new Error("This coupon is only valid for " + scope + ".");
    const offerType = String(r[col.offer_type]).toLowerCase();
    const num = Number(String(r[col.offer_value]).replace(/[^0-9.]/g, ""));
    if (!isFinite(num) || num <= 0) throw new Error("This coupon is not set up correctly. Please contact MyPod.");
    const percent = offerType.indexOf("percent") >= 0;
    return { code: String(r[col.coupon_code]).trim().toUpperCase(), type: percent ? "percent" : "fixed", value: num, rowNumber: i + 1,
             label: percent ? num + "% off applied" : "Flyer offer ₹" + num + " applied" };
  }
  badCouponTry_();
  throw new Error("Coupon code not found.");
}

function normalizeCouponCode_(value) {
  const raw = String(value || "").trim().toUpperCase();
  const compact = raw.replace(/\bAUTO\b|\bDASH\b/g, "").replace(/[\s‐‑‒–—-]+/g, "");
  const match = /^MP(CER|PPF|CW|BW)(\d{1,3})$/.exec(compact);
  if (!match) return raw;
  const number = Number(match[2]);
  return number >= 1 && number <= 300 ? "MP-" + match[1] + "-" + ("00" + number).slice(-3) : raw;
}

// Allows only a limited number of wrong codes per 15 minutes, because codes are numbered in order and could be guessed.
function badCouponTry_() {
  const cache = CacheService.getScriptCache(), key = "bad_coupon";
  const n = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(n), 900);
  if (n > MAX_BAD_COUPON_TRIES) throw new Error("Too many wrong coupon attempts. Please wait a few minutes and try again.");
}

// "percent": percent off the list price. "fixed": the coupon sets the price (for example Rs 299), so the discount is the difference.
function discountFor_(coupon, price) {
  if (price === null) return 0;
  const d = coupon.type === "percent" ? Math.round(price * coupon.value / 100) : price - coupon.value;
  return Math.min(price, Math.max(0, d));
}

// Marks the coupon as used. Called inside the booking lock, right after the booking row is saved.
function redeemCoupon_(coupon, bookingId, name, phone) {
  const data = loadCoupons_(), col = data.col, sh = data.sheet, row = coupon.rowNumber;
  const set = (h, v) => sh.getRange(row, col[h] + 1).setValue(v);
  const now = Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
  set("status", "Redeemed");
  set("used_count", Number(sh.getRange(row, col.used_count + 1).getValue() || 0) + 1);
  set("customer_name", name); sh.getRange(row, col.customer_phone + 1).setNumberFormat("@").setValue(phone); set("booking_id", bookingId); set("redeemed_at", now);
}

// A cancelled booking gives its coupon back (only if that coupon is still tied to this booking).
function releaseCoupon_(code, bookingId) {
  try {
    const c = String(code || "").trim().toUpperCase(); if (!c) return;
    const data = loadCoupons_(), col = data.col, sh = data.sheet;
    for (let i = 1; i < data.rows.length; i++) {
      if (String(data.rows[i][col.coupon_code]).trim().toUpperCase() !== c) continue;
      if (String(data.rows[i][col.booking_id]) !== String(bookingId)) return;
      const row = i + 1, set = (h, v) => sh.getRange(row, col[h] + 1).setValue(v);
      set("status", "Active"); set("used_count", 0);
      ["customer_name", "customer_phone", "booking_id", "redeemed_at"].forEach(h => set(h, ""));
      set("notes", "Released after cancellation of " + bookingId);
      return;
    }
  } catch (e) { /* never block a cancellation on this */ }
}

// Doorstep only: find the customer's coordinates (GPS from the site, or by geocoding the typed address), confirm they are inside
// the Newtown service area, and read the pincode and locality from Google Maps. Returns { area, map }.
function resolveArea_(b) {
  const given = b.lat !== null && b.lat !== undefined && b.lat !== "" && b.lng !== null && b.lng !== undefined && b.lng !== "";
  let lat = given ? Number(b.lat) : NaN, lng = given ? Number(b.lng) : NaN, pin = "", place = "";
  try {
    const gc = Maps.newGeocoder().setRegion("in");
    if (!isFinite(lat) || !isFinite(lng)) {
      const f = gc.geocode(String(b.address) + ", Kolkata");
      if (f.status === "OK" && f.results.length) { lat = f.results[0].geometry.location.lat; lng = f.results[0].geometry.location.lng; }
    }
    if (isFinite(lat) && isFinite(lng)) {
      const r = gc.reverseGeocode(lat, lng);
      if (r.status === "OK" && r.results.length) {
        const comps = r.results[0].address_components || [];
        const pick = t => (comps.filter(c => c.types.indexOf(t) >= 0)[0] || {}).long_name || "";
        pin = pick("postal_code"); place = pick("sublocality_level_1") || pick("sublocality") || pick("locality");
      }
    }
  } catch (e) { /* Maps unavailable: fall back to the checks below */ }
  const hasPoint = isFinite(lat) && isFinite(lng);
  const typedOk = NEWTOWN_RE.test(String(b.address)) || NEWTOWN_RE.test(pin);
  if (given && hasPoint ? !inArea_(lat, lng) : (hasPoint ? !(inArea_(lat, lng) || typedOk) : !typedOk))
    throw new Error("Doorstep service covers Newtown and one nearby area only. Please choose drive-through or an address inside our service area.");
  return { area: [pin, place].filter(String).join(" · ") || "Newtown", map: hasPoint ? "https://www.google.com/maps?q=" + lat + "," + lng : "" };
}

function inArea_(lat, lng) {
  const rad = Math.PI / 180;
  if (EXTRA_AREAS.some(c => {
    const a = Math.pow(Math.sin((lat - c.lat) * rad / 2), 2) + Math.cos(lat * rad) * Math.cos(c.lat * rad) * Math.pow(Math.sin((lng - c.lng) * rad / 2), 2);
    return 2 * 6371000 * Math.asin(Math.sqrt(a)) <= c.radius;
  })) return true;
  let inside = false;
  for (let i = 0, j = SERVICE_AREA.length - 1; i < SERVICE_AREA.length; j = i++) {
    const yi = SERVICE_AREA[i][0], xi = SERVICE_AREA[i][1], yj = SERVICE_AREA[j][0], xj = SERVICE_AREA[j][1];
    if ((yi > lat) !== (yj > lat) && lng < (xj - xi) * (lat - yi) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
}

function rateLimit_(phone) {
  const cache = CacheService.getScriptCache(), key = "rl_" + phone;
  const n = Number(cache.get(key) || 0);
  if (n >= MAX_BOOKINGS_PER_PHONE_PER_HOUR) throw new Error("Too many bookings from this number. Please call us to book more.");
  cache.put(key, String(n + 1), 3600);
}

function notifyOwner_(subject, o) {
  if (!OWNER_EMAIL) return;
  const lines = Object.keys(o).map(k => k + ": " + o[k]).join("\n");
  try { MailApp.sendEmail(OWNER_EMAIL, "[MyPod] " + subject, lines); } catch (e) { /* never block a booking on email */ }
}

function esc_(v) {
  return String(v == null ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function prettyDate_(ymd) {
  const d = new Date(String(ymd) + "T12:00:00+05:30");
  return Utilities.formatDate(d, TIME_ZONE, "EEEE, d MMMM yyyy");
}

// Sends the customer an HTML email. A failed email never blocks or undoes the booking.
function sendCustomerEmail_(kind, o) {
  if (!SEND_CUSTOMER_EMAIL || !o["Email"]) return;
  try {
    const id = o["Booking ID"], first = String(o["Customer Name"]).trim().split(/\s+/)[0];
    const total = o["Total (INR)"] === "" ? "Final quote confirmed by MyPod" : "₹" + Number(o["Total (INR)"]).toLocaleString("en-IN");
    const cancelled = kind === "cancelled";
    const rows = [
      ["Booking ID", id],
      ["Service", o["Service"] + " (" + o["Vehicle Type"] + ")"],
      ["Date", prettyDate_(normalizeDate_(o["Date"]))],
      ["Time", normalizeTime_(o["Time"])],
      [o["Service Mode"] === "Doorstep" ? "Doorstep address" : "Drive-through at", o["Address"]],
      ["Amount", total]
    ];
    if (o["Coupon Code"]) rows.splice(5, 0, ["Coupon", o["Coupon Code"] + " (− ₹" + Number(o["Discount (INR)"] || 0).toLocaleString("en-IN") + ")"]);
    const tr = rows.map(r => '<tr><td style="padding:9px 0;color:#8a8a8a;font-size:13px;width:38%;vertical-align:top">' + esc_(r[0]) +
      '</td><td style="padding:9px 0;color:#f4f4f4;font-size:14px;font-weight:600">' + esc_(r[1]) + '</td></tr>').join("");
    const wa = "https://wa.me/" + SUPPORT_WHATSAPP + "?text=" + encodeURIComponent("Hi MyPod, about my booking " + id);
    const manage = SITE_URL ? '<a href="' + esc_(SITE_URL) + '#manage" style="color:#ffd51b">manage or cancel your booking</a>' : "";
    const heading = cancelled ? "Your booking is cancelled" : "Your booking is confirmed";
    const intro = cancelled
      ? "Hi " + esc_(first) + ", your MyPod booking has been cancelled and the slot is released. You are welcome to book again any time."
      : "Hi " + esc_(first) + ", thank you for choosing MyPod. Your slot is reserved and our team will be there on time.";
    const next = cancelled ? "" :
      '<p style="margin:22px 0 6px;color:#f4f4f4;font-size:15px;font-weight:700">What happens next</p>' +
      '<ul style="margin:0;padding-left:18px;color:#cfcfcf;font-size:13px;line-height:1.7">' +
      (o["Service Mode"] === "Doorstep"
        ? '<li>Our crew arrives at your address at the booked time. Please keep the car accessible.</li><li>Keep your phone nearby in case the crew needs to reach you.</li>'
        : '<li>Drive in to ' + esc_(STUDIO_ADDRESS) + ' a few minutes before your slot.</li>') +
      '<li>Need to change something? Reply to this email or message us on WhatsApp' + (manage ? ", or " + manage : "") + '.</li></ul>';
    const mapBtn = (!cancelled && o["Map Link"]) ? '<a href="' + esc_(o["Map Link"]) + '" style="display:inline-block;margin:6px 8px 0 0;padding:11px 18px;border:1px solid #555;border-radius:24px;color:#fff;text-decoration:none;font-size:13px;font-weight:700">View location</a>' : "";
    const html =
      '<div style="background:#050505;padding:24px 12px;font-family:Arial,Helvetica,sans-serif">' +
      '<table role="presentation" width="100%" style="max-width:560px;margin:0 auto;background:#0e0e0e;border:1px solid #2a2a2a;border-radius:16px" cellpadding="0" cellspacing="0"><tr><td style="padding:28px 26px">' +
      '<div style="font-size:26px;font-weight:900;letter-spacing:-1px;color:#fff">MY<span style="color:#ffd51b">POD</span></div>' +
      '<div style="font-size:10px;letter-spacing:3px;color:#9a9a9a;margin-top:4px">' + (cancelled ? "BOOKING UPDATE" : "BOOKING CONFIRMATION") + '</div>' +
      '<h1 style="margin:22px 0 8px;font-size:24px;line-height:1.25;color:#ffd51b">' + heading + '</h1>' +
      '<p style="margin:0 0 18px;color:#d8d8d8;font-size:14px;line-height:1.6">' + intro + '</p>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #2a2a2a;border-bottom:1px solid #2a2a2a">' + tr + '</table>' +
      next +
      '<p style="margin:22px 0 0">' + mapBtn +
      '<a href="' + wa + '" style="display:inline-block;margin-top:6px;padding:11px 18px;background:#ffd51b;border-radius:24px;color:#111;text-decoration:none;font-size:13px;font-weight:800">Chat on WhatsApp</a></p>' +
      '<p style="margin:26px 0 0;color:#7d7d7d;font-size:11px;line-height:1.6">' + esc_(BRAND_NAME) + ' · Newtown, Kolkata<br>You received this email because a booking was made with this address.</p>' +
      '</td></tr></table></div>';
    const text = heading + "\n\n" + rows.map(r => r[0] + ": " + r[1]).join("\n") + "\n\nWhatsApp: " + wa;
    const opts = { name: BRAND_NAME, htmlBody: html };
    if (OWNER_EMAIL) opts.replyTo = OWNER_EMAIL;
    MailApp.sendEmail(o["Email"], (cancelled ? "Booking cancelled · " : "Booking confirmed · ") + id + " · " + BRAND_NAME, text, opts);
  } catch (e) { /* ignore: booking is already saved */ }
}

// Run this once from the editor to preview the email in your own inbox.
function testCustomerEmail() {
  if (!OWNER_EMAIL) throw new Error("Set OWNER_EMAIL first.");
  sendCustomerEmail_("confirmed", { "Booking ID": "MP-TEST-0001", "Customer Name": "Test Customer", "Email": OWNER_EMAIL, "Service": "Promotional Car Wash",
    "Vehicle Type": "Car", "Date": Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd"), "Time": "10:00 AM", "Service Mode": "Doorstep",
    "Address": "Flat 4B, Action Area I, Newtown", "Coupon Code": "", "Discount (INR)": 0, "Total (INR)": 299, "Map Link": "" });
  return "Test email sent to " + OWNER_EMAIL;
}

function isActive_(status) { return ["Confirmed", "Pending"].indexOf(String(status)) >= 0; }
function checkMode_(m) { if (!SLOT_CAPACITY[m]) throw new Error("Invalid service mode."); }

function checkDate_(d) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d))) throw new Error("Invalid booking date.");
  const today = Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd");
  if (String(d) < today) throw new Error("Booking date cannot be in the past.");
}

function requireFields_(b, keys) {
  keys.forEach(k => { if (!String(b[k] == null ? "" : b[k]).trim()) throw new Error("Missing required field: " + k); });
}

function normalizeEmail_(v) {
  const e = String(v || "").trim().toLowerCase();
  if (!/^[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}$/.test(e) || e.length > 200) throw new Error("Enter a valid email address.");
  return e;
}

function normalizePhone_(v) {
  const digits = String(v || "").replace(/\D/g, "").slice(-10);
  if (!/^[6-9]\d{9}$/.test(digits)) throw new Error("Enter a valid 10-digit Indian mobile number.");
  return digits;
}

function normalizeDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TIME_ZONE, "yyyy-MM-dd");
  const serial = typeof v === "number" ? v : (/^\d{1,5}(?:\.\d+)?$/.test(String(v || "")) ? Number(v) : NaN);
  if (isFinite(serial) && serial >= 20000 && serial <= 100000)
    return Utilities.formatDate(new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 864e5), TIME_ZONE, "yyyy-MM-dd");
  return String(v || "").slice(0, 10);
}

function normalizeTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TIME_ZONE, "h:mm a");   // Sheets auto-converted a time string
  return String(v || "").trim();
}

function clean_(value) {
  // Stop spreadsheet formula injection from customer-entered text.
  const s = String(value == null ? "" : value).trim().slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
