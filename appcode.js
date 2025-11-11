const SHEET_ID = "1f6WHZz1y_7OblNPLGSw3cSe8UY97vqmD9UzBNQTxS5g";
const SHEET_NAME = "FormResponses";

// LINE Messaging API: ตั้งค่า Channel Access Token ตรงนี้ (ใส่ค่าโดยตรงในโค้ด)
// หมายเหตุ: เปลี่ยนสตริงด้านล่างเป็น Channel Access Token ของบอทคุณ
const LINE_CHANNEL_ACCESS_TOKEN =
  "bwfEzqJqn1bIkHKdX7lPgPF33+4HsYMlL/gIMGLrhkvo053s1Ig8JEzu1lcO2Ci5bEDESBSWkKcLtPIY+hA74lQmCuXzQDNj1/YDegnwckwjmqh9y5BV0Tz2tWtYBHbzCrURW6mMTZBv/to93me8YAdB04t89/1O/w1cDnyilFU=";
const LINE_REPLY_URL = "https://api.line.me/v2/bot/message/reply";
const LINE_PROFILE_URL = "https://api.line.me/v2/bot/profile/";

// เรียกด้วย GET (เช่น เปิดลิงก์ใน browser ตรงๆ)
// เอาไว้เช็คว่า Web App ทำงาน + ตั้ง permission ถูก
function doGet(e) {
  const params = e && e.parameter;

  // ทดสอบดู JSON ของ Flex Message
  if (params && params.action === "testFlex") {
    const flexMsg = testFlexMessage();
    return ContentService.createTextOutput(JSON.stringify(flexMsg, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(
    "OK: Web App is running. Use POST to submit data."
  ).setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    // 1) ตรวจจับ LINE Webhook ก่อน (application/json ที่มี events)
    if (isLineWebhook(e)) {
      handleLineWebhook(e);
      return ContentService.createTextOutput("OK").setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    // 2) โหมดฟอร์ม (หน้าเว็บ) เดิม - ต้องมี action parameter
    const req = parseRequest(e);

    // ถ้าไม่มี action หรือ lineUserId แสดงว่าไม่ใช่คำขอจากฟอร์ม (อาจเป็นข้อความจากแชท)
    if (!req.action && !req.lineUserId) {
      return jsonOutput({
        status: "ignored",
        message: "Not a form submission",
      });
    }

    // เตรียมชีต + หัวตาราง
    const sheet = ensureSheet();

    // เช็คซ้ำตาม LINE UserID
    if (req.action === "check") {
      const exists =
        !!req.lineUserId && findByLineUserId(sheet, req.lineUserId) !== -1;
      return jsonOutput({ status: exists ? "exists" : "not_found", exists });
    }

    // สมัครจริง: ต้องมี lineUserId
    if (!req.lineUserId) {
      return jsonOutput({ status: "error", message: "Missing lineUserId" });
    }

    // ถ้ามีอยู่แล้ว ไม่บันทึกซ้ำ
    const rowIndex = findByLineUserId(sheet, req.lineUserId);
    if (rowIndex !== -1) {
      return jsonOutput({
        status: "exists",
        exists: true,
        message: "Duplicate lineUserId",
      });
    }

    // บันทึกแถวใหม่
    sheet.appendRow([
      new Date(),
      req.lineUserId || "",
      req.lineDisplayName || "",
      req.companyType || "",
      req.company || "",
      req.fullName || "",
      req.nickname || "",
      req.phone || "",
      req.email || "",
      req.address || "",
      req.userAgent || "",
    ]);

    return jsonOutput({ status: "success", message: "Saved to Google Sheet" });
  } catch (error) {
    console.error("doPost error:", error);
    return jsonOutput({ status: "error", message: error.toString() });
  }
}

// ---------- LINE Webhook Handlers ----------

function isLineWebhook(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return false;
    const contentType = (e.postData.type || "").toLowerCase();
    if (contentType.indexOf("application/json") === -1) return false;
    const body = JSON.parse(e.postData.contents);
    return body && Array.isArray(body.events);
  } catch (_) {
    return false;
  }
}

function handleLineWebhook(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents || "{}");
  } catch (err) {
    console.error("Invalid JSON webhook:", err);
    return;
  }

  const events = body.events || [];
  events.forEach(function (event) {
    try {
      if (
        event.type === "message" &&
        event.message &&
        event.message.type === "text"
      ) {
        const text = String(event.message.text || "").trim();
        const norm = normalizeThai(text);
        // คีย์เวิร์ดที่ใช้เรียกดูคะแนน
        const keywords = [
          "point",
          "points",
          "คะแนน",
          "คะแนนของฉัน",
          "พอยท์",
          "พอยต์",
          "แต้ม",
        ];
        const matched = keywords.some(function (k) {
          return norm.indexOf(normalizeThai(k)) !== -1;
        });

        if (matched) {
          replyUserPoints(event);
          return;
        }

        // ข้อความอื่นๆ: ตอบกลับข้อความเดียวกัน (echo)
        const replyToken = event.replyToken;
        if (replyToken) {
          const messages = [{ type: "text", text: text }];
          lineReply(replyToken, messages);
        }
      }
      // ไม่จัดการอีเวนต์อื่น ๆ
    } catch (evErr) {
      console.error("handle event error:", evErr);
    }
  });
}

function replyUserPoints(event) {
  const userId = (event && event.source && event.source.userId) || "";
  const replyToken = event.replyToken;
  if (!userId || !replyToken) return;

  const sheet = ensureSheet();
  const rowIndex = findByLineUserId(sheet, userId);

  let displayName = "";
  try {
    // พยายามอ่านชื่อจากชีตก่อน (คอลัมน์ 3: LINE Name)
    if (rowIndex !== -1) {
      const vals = sheet
        .getRange(rowIndex, 1, 1, sheet.getLastColumn())
        .getValues()[0];
      displayName = String(vals[2] || "");
    }
    if (!displayName) {
      // fallback: ดึงโปรไฟล์จาก LINE API
      const prof = getLineProfile(userId);
      displayName = (prof && prof.displayName) || "";
    }
  } catch (_) { }

  const orderTotal = getOrderTotalForUser(sheet, rowIndex);
  const points = Math.floor(orderTotal / 100);

  const flexBubble = buildFlexPointsBubble({
    displayName: displayName,
    orderTotal: orderTotal,
    points: points,
  });

  const messages = [
    {
      type: "flex",
      altText: "คะแนนสะสมของคุณ: " + points + " POINTS",
      contents: flexBubble,
    },
  ];
  lineReply(replyToken, messages);
}

function buildPointsTextMessage(data) {
  const name = data.displayName || "สมาชิก";
  const total = data.orderTotal || 0;
  const points = data.points || 0;
  const totalFmt = formatCurrency(total);
  var lines = [];
  lines.push("⭐ คะแนนสะสมของคุณ");
  lines.push("ชื่อ: " + name);
  lines.push("ยอดสั่งซื้อรวม: ฿" + totalFmt);
  lines.push("คะแนน: " + String(points) + " POINTS");
  lines.push("อัตราแลกแต้ม: 100 บาท = 1 POINT");
  return lines.join("\n");
}

function getLineProfile(userId) {
  if (!LINE_CHANNEL_ACCESS_TOKEN) return null;
  try {
    const res = UrlFetchApp.fetch(
      LINE_PROFILE_URL + encodeURIComponent(userId),
      {
        method: "get",
        headers: { Authorization: "Bearer " + LINE_CHANNEL_ACCESS_TOKEN },
        muteHttpExceptions: true,
      }
    );
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      return JSON.parse(res.getContentText() || "{}");
    }
  } catch (err) {
    console.warn("getLineProfile error:", err);
  }
  return null;
}

function lineReply(replyToken, messages) {
  if (!LINE_CHANNEL_ACCESS_TOKEN) {
    console.warn("Missing LINE_CHANNEL_ACCESS_TOKEN: skip reply");
    return;
  }
  try {
    const payload = JSON.stringify({
      replyToken: replyToken,
      messages: messages,
    });
    const res = UrlFetchApp.fetch(LINE_REPLY_URL, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + LINE_CHANNEL_ACCESS_TOKEN },
      payload: payload,
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      console.error("LINE reply failed:", code, res.getContentText());
    }
  } catch (err) {
    console.error("lineReply error:", err);
  }
}

function getOrderTotalForUser(sheet, rowIndex) {
  if (!sheet || rowIndex === -1) return 0;
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  // พยายามหาโดยชื่อหัวคอลัมน์ก่อน รองรับสะกดทั้ง "ยอดสั่งซื้อ" และ "ยอดสั่งซ์้อ"
  const possibleHeaders = [
    "ยอดสั่งซื้อ",
    "ยอดสั่งซ์้อ",
    "ยอดซื้อ",
    "ยอดออเดอร์",
  ];
  let colIndex = -1; // 1-based
  for (var i = 0; i < header.length; i++) {
    var h = String(header[i] || "").trim();
    if (possibleHeaders.indexOf(h) !== -1) {
      colIndex = i + 1;
      break;
    }
  }
  // ไม่เจอชื่อหัว: พยายามอ่านจากคอลัมน์ K (11) ตามที่แจ้ง
  // หากมีโครงสร้างแผ่นงานที่ถูกขยับ (เช่น เพิ่มคอลัมน์ก่อนหน้า) ให้ลองอ่าน L (12) เป็นแผนสำรอง
  if (colIndex === -1) {
    try {
      const vK = sheet.getRange(rowIndex, 11).getValue();
      const nK = toNumber(vK);
      if (nK > 0) return nK;
      const vL = sheet.getRange(rowIndex, 12).getValue();
      const nL = toNumber(vL);
      return nL > 0 ? nL : 0;
    } catch (err) {
      console.warn("getOrderTotalForUser read fallback error:", err);
      return 0;
    }
  }

  try {
    const val = sheet.getRange(rowIndex, colIndex).getValue();
    return toNumber(val);
  } catch (err) {
    console.warn("getOrderTotalForUser read error:", err);
    return 0;
  }
}

function toNumber(value) {
  if (typeof value === "number") return value;
  var s = String(value || "").replace(/[,฿\s]/g, "");
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function normalizeThai(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]/g, "") // ลบวรรณยุกต์เพื่อเทียบง่าย
    .trim();
}

function buildFlexPointsBubble(data) {
  const name = data.displayName || "สมาชิก";
  const total = data.orderTotal || 0;
  const points = data.points || 0;
  const totalFmt = formatCurrency(total);

  return {
    type: "bubble",
    size: "mega",
    hero: {
      type: "box",
      layout: "vertical",
      contents: [
        {
          type: "box",
          layout: "vertical",
          contents: [],
          position: "absolute",
          width: "100%",
          height: "100%",
          background: {
            type: "linearGradient",
            angle: "180deg",
            endColor: "#00000099",
            startColor: "#00000010"
          }
        },
        {
          type: "box",
          layout: "vertical",
          contents: [
            {
              type: "text",
              text: "บัตรสมาชิก",
              color: "#FFFFFF",
              size: "sm",
              weight: "bold"
            },
            {
              type: "text",
              text: name,
              color: "#FFFFFF",
              size: "xl",
              weight: "bold",
              margin: "xs"
            }
          ],
          position: "absolute",
          offsetBottom: "20px",
          offsetStart: "20px"
        }
      ],
      paddingAll: "0px"
    },
    body: {
      type: "box",
      layout: "vertical",
      contents: [
        {
          type: "box",
          layout: "vertical",
          contents: [
            {
              type: "image",
              url: "https://bam-samada.github.io/line-member-form/assets/logo/1030766.jpg",
              size: "full",
              aspectMode: "cover",
              gravity: "center",
              align: "center",
              aspectRatio: "1.7:2"
            }
          ],
          cornerRadius: "15px"
        },
        {
          type: "box",
          layout: "horizontal",
          contents: [
            {
              type: "image",
              url: "https://bam-samada.github.io/line-member-form/assets/logo/CI%20LOGO-06.png",
              size: "xxs",
              aspectMode: "fit",
              flex: 0
            },
            {
              type: "box",
              layout: "baseline",
              contents: [
                {
                  type: "text",
                  text: "คะแนนสะสมของคุณ",
                  size: "sm",
                  color: "#78716C",
                  align: "center",
                  weight: "bold",
                  offsetTop: "10px"
                }
              ],
              flex: 1
            },
            {
              type: "image",
              url: "https://bam-samada.github.io/line-member-form/assets/logo/S-MAN%20%E0%B8%8A%E0%B8%B9%E0%B8%99%E0%B8%B4%E0%B9%89%E0%B8%A7%E0%B8%8A%E0%B8%B5%E0%B9%89%E0%B8%82%E0%B8%B6%E0%B9%89%E0%B8%99%E0%B8%82%E0%B9%89%E0%B8%B2%E0%B8%87%E0%B8%9A%E0%B8%99%201.jpg",
              size: "xxs",
              aspectMode: "fit",
              flex: 0
            }
          ],
          position: "relative",
          paddingTop: "10px"
        },
        {
          type: "box",
          layout: "vertical",
          contents: [
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "box",
                  layout: "vertical",
                  contents: [
                    {
                      type: "text",
                      text: "💰",
                      size: "xl"
                    }
                  ],
                  flex: 0,
                  justifyContent: "center",
                  alignItems: "center",
                  paddingAll: "8px",
                  backgroundColor: "#FEF3C7",
                  cornerRadius: "12px",
                  margin: "none"
                },
                {
                  type: "box",
                  layout: "vertical",
                  contents: [
                    {
                      type: "text",
                      text: "คะแนน",
                      size: "xs",
                      color: "#9CA3AF"
                    },
                    {
                      type: "text",
                      text: String(points),
                      size: "xl",
                      weight: "bold",
                      color: "#1F2937",
                      margin: "xs"
                    }
                  ],
                  flex: 1,
                  margin: "md"
                }
              ]
            },
            {
              type: "separator",
              margin: "lg",
              color: "#E5E7EB"
            },
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "box",
                  layout: "vertical",
                  contents: [
                    {
                      type: "text",
                      text: "💰",
                      size: "xl"
                    }
                  ],
                  flex: 0,
                  justifyContent: "center",
                  alignItems: "center",
                  paddingAll: "8px",
                  backgroundColor: "#14C08199",
                  cornerRadius: "12px"
                },
                {
                  type: "box",
                  layout: "vertical",
                  contents: [
                    {
                      type: "text",
                      text: "ยอดสั่งซื้อรวม",
                      size: "xs",
                      color: "#9CA3AF"
                    },
                    {
                      type: "text",
                      text: "฿" + totalFmt,
                      size: "lg",
                      weight: "bold",
                      color: "#1F2937",
                      margin: "xs"
                    }
                  ],
                  flex: 1,
                  margin: "md"
                }
              ],
              margin: "lg"
            },
            {
              type: "separator",
              margin: "lg",
              color: "#E5E7EB"
            },
            {
              type: "box",
              layout: "horizontal",
              contents: [
                {
                  type: "box",
                  layout: "vertical",
                  contents: [
                    {
                      type: "text",
                      text: "⚡",
                      size: "xl"
                    }
                  ],
                  flex: 0,
                  justifyContent: "center",
                  alignItems: "center",
                  paddingAll: "8px",
                  backgroundColor: "#DBEAFE",
                  cornerRadius: "12px"
                },
                {
                  type: "box",
                  layout: "vertical",
                  contents: [
                    {
                      type: "text",
                      text: "อัตราแลกแต้ม",
                      size: "xs",
                      color: "#9CA3AF"
                    },
                    {
                      type: "text",
                      text: "100฿ = 1 Point",
                      size: "lg",
                      weight: "bold",
                      color: "#1F2937",
                      margin: "xs"
                    }
                  ],
                  flex: 1,
                  margin: "md"
                }
              ],
              margin: "lg"
            }
          ],
          paddingAll: "20px",
          backgroundColor: "#FFFFFF",
          cornerRadius: "20px",
          borderWidth: "2px",
          borderColor: "#1B7DE799",
          margin: "xl"
        },
        {
          type: "box",
          layout: "vertical",
          contents: [
            {
              type: "text",
              text: "สะสมแต้มได้ทุกการซื้อ",
              size: "xs",
              color: "#6B7280",
              align: "center"
            }
          ],
          paddingAll: "12px",
          backgroundColor: "#F9FAFB",
          cornerRadius: "16px",
          margin: "lg"
        }
      ],
      spacing: "none"
    }
  };
}

function formatCurrency(n) {
  try {
    return Number(n || 0).toLocaleString("th-TH", { maximumFractionDigits: 2 });
  } catch (_) {
    return String(n || 0);
  }
}

// แยก parse request ให้รองรับทั้ง JSON / text/plain(JSON) / x-www-form-urlencoded
function parseRequest(e) {
  if (!e) return {};
  const p = e.parameter || {};
  const contentType = (e.postData && e.postData.type) || "";
  const body = (e.postData && e.postData.contents) || "";

  let fromBody = {};
  try {
    if (
      contentType.indexOf("application/json") !== -1 ||
      contentType.indexOf("text/plain") !== -1
    ) {
      fromBody = body ? JSON.parse(body) : {};
    }
  } catch (_) {
    // ถ้า parse ไม่ได้ จะใช้ค่าจาก form parameters แทน
  }

  const get = (k) => (fromBody[k] != null ? fromBody[k] : p[k]) || "";
  return {
    action: (get("action") || "").toString().trim() || "submit",
    lineUserId: (get("lineUserId") || "").toString().trim(),
    lineDisplayName: (get("lineDisplayName") || "").toString().trim(),
    companyType: (get("companyType") || "").toString().trim(),
    company: (get("company") || "").toString().trim(),
    fullName: (get("fullName") || "").toString().trim(),
    nickname: (get("nickname") || "").toString().trim(),
    phone: (get("phone") || "").toString().trim(),
    email: (get("email") || "").toString().trim(),
    address: (get("address") || "").toString().trim(),
    userAgent: (get("userAgent") || "").toString().trim(),
  };
}

function ensureSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp",
      "LINE UserID",
      "LINE Name",
      "Company Type",
      "Company",
      "Full Name",
      "Nickname",
      "Phone",
      "Email",
      "Address",
      "User Agent",
    ]);
  }
  ensureCompanyTypeColumn(sheet);
  return sheet;
}

// Ensure the sheet has a dedicated column for Company Type at column 4
function ensureCompanyTypeColumn(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (header.indexOf("Company Type") === -1) {
    // Insert after column 3 (after LINE Name)
    sheet.insertColumnAfter(3);
    sheet.getRange(1, 4).setValue("Company Type");
  }
}

// หา row index ของ LINE UserID (ไม่รวม header) ถ้าไม่พบ return -1
function findByLineUserId(sheet, lineUserId) {
  if (!lineUserId) return -1;
  const values = sheet.getDataRange().getValues();
  // คอลัมน์ที่ 2 (index 1) คือ LINE UserID ตามหัวข้อที่ตั้งไว้
  for (let i = 1; i < values.length; i++) {
    if ((values[i][1] || "").toString().trim() === lineUserId) {
      return i + 1; // row index (1-based)
    }
  }
  return -1;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ฟังก์ชันทดสอบ: ใช้สำหรับดู JSON ของ Flex Message
// เรียกใช้ผ่าน GET parameter ?action=testFlex
// เช่น: https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?action=testFlex
function testFlexMessage() {
  const sampleData = {
    displayName: "ทดสอบ ระบบ",
    orderTotal: 12345.67,
    points: 123
  };

  const flexBubble = buildFlexPointsBubble(sampleData);

  return {
    type: "flex",
    altText: "คะแนนสะสมของคุณ: " + sampleData.points + " POINTS",
    contents: flexBubble
  };
}
