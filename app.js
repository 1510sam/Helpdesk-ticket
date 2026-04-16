const SHEET_NAME = "IncidentLog";
const SPREADSHEET_ID = "1xkPeI21txXHuFUpqZhYN8lEgHXHa_I1xlsqAMQAHusI";
const SHEET_USER = "users";
const SHEET_NS = "ns_lilichnv";
const DRIVE_FOLDER_ID = "1otAxAX3cPs4lVr4W_FqSslI6SEf2XS7b";

// --- CẤU HÌNH DANH SÁCH DROPDOWN VÀ MÀU SẮC ---
const OPTIONS = {
  LOAI_SU_CO: {
    HIS: { bg: "#d9ead3", fg: "#274e13" },
    BIS: { bg: "#cfe2f3", fg: "#073763" },
    RIS: { bg: "#d0e0e3", fg: "#134f5c" },
    LIS: { bg: "#c9daf8", fg: "#1c4587" },
  },

  NGUOI_TIEP_NHAN: {
    Tuấn: { bg: "#cfe2f3", fg: "#073763" },
    Hùng: { bg: "#d9ead3", fg: "#274e13" },
    Khoa: { bg: "#ead1dc", fg: "#741b47" },
    Huy: { bg: "#fff2cc", fg: "#7f6000" },
  },

  TRANG_THAI: {
    "Chưa xử lý": { bg: "#f3f4f6", fg: "#374151" },
    "Chờ xử lý": { bg: "#fff2cc", fg: "#7f6000" },
    "Đang xử lý": { bg: "#cfe2f3", fg: "#073763" },
    "Hoàn tất": { bg: "#d9ead3", fg: "#274e13" },
    Hủy: { bg: "#f4cccc", fg: "#990000" },
  },
};

const DEFAULT_STATUS = {
  NGUOI_TIEP_NHAN: "Chưa xét",
  MUC_DO: "Chưa xét",
  TRANG_THAI: "Chưa xử lý",
};

function calculateDiffMinutes(startStr, endStr) {
  if (!startStr || !endStr) return 0;
  const start = new Date(startStr);
  const end = new Date(endStr);
  // Check date hợp lệ
  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
    return 0;
  }
  const diffMs = end.getTime() - start.getTime();
  // Không cho âm
  if (diffMs <= 0) return 0;
  // Làm tròn xuống để chuẩn SLA
  return Math.floor(diffMs / 60000);
}

function parseVNDateTime(str) {
  if (!str) return null;

  // Format: DD/MM/YYYY HH:mm
  const [datePart, timePart] = str.split(" ");
  if (!datePart || !timePart) return null;

  const [d, m, y] = datePart.split("/");
  const [h, min] = timePart.split(":");

  return new Date(Number(y), Number(m) - 1, Number(d), Number(h), Number(min));
}

function checkLogin(username, password, role, phone) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const isQuickLogin = password === "QUICK_LOGIN_NVBV";

  if (isQuickLogin) {
    const sheetNS = ss.getSheetByName("ns_lilichnv");
    if (!sheetNS) {
      console.log("Sheet ns_lilichnv không tồn tại!");
      return {
        success: false,
        message: "Lỗi: Không tìm thấy sheet ns_lilichnv!",
      };
    }

    const dataNS = sheetNS.getDataRange().getValues();

    for (let i = 2; i < dataNS.length; i++) {
      const maNVFromSheet = String(dataNS[i][1]).trim(); // Cột B
      const hoTenFromSheet = String(dataNS[i][6]).trim(); // Cột G
      const maBPFromSheet = String(dataNS[i][2]).trim(); // Cột C
      const soDienThoaiFromSheet = String(dataNS[i][25]).trim(); // Cột Z

      if (
        maNVFromSheet.toLowerCase() === username.toLowerCase() &&
        soDienThoaiFromSheet === phone
      ) {
        return {
          success: true,
          targetPage: "create",
          user: {
            username: maNVFromSheet,
            fullName: hoTenFromSheet,
            department: maBPFromSheet,
            phone: soDienThoaiFromSheet,
            role: "khach",
          },
        };
      }
    }
    return {
      success: false,
      message: "Nhập sai mã nhân viên hoặc số điện thoại!",
    };
  }

  const sheetUser = ss.getSheetByName(SHEET_USER);
  if (!sheetUser)
    return { success: false, message: "Lỗi: Không tìm thấy sheet tài khoản!" };

  const dataUser = sheetUser.getDataRange().getValues();

  for (let i = 1; i < dataUser.length; i++) {
    const usernameSheet = String(dataUser[i][0]).trim();
    const passwordSheet = String(dataUser[i][1]).trim();
    const roleSheet = String(dataUser[i][2]).trim();
    const fullNameSheet = String(dataUser[i][3] || "").trim();

    if (
      usernameSheet.toLowerCase() === username.toLowerCase() &&
      passwordSheet === password &&
      roleSheet.toLowerCase() === role.toLowerCase()
    ) {
      let targetPage = "view"; // default
      if (roleSheet.toLowerCase() === "it") {
        targetPage = "admin"; // ← Sửa ở đây: "admin" để map sang 'form-view' trong getPageContent
      }

      return {
        success: true,
        targetPage: targetPage,
        user: {
          username: usernameSheet,
          role: roleSheet,
          fullName: fullNameSheet,
        },
      };
    }
  }
  return { success: false, message: "Tài khoản hoặc mật khẩu không đúng!" };
}

function getPageContent(targetPage, userJson) {
  console.log("getPageContent called with targetPage:", targetPage);
  const user = JSON.parse(userJson);
  let fileName = "";

  if (targetPage === "create") fileName = "form-create";
  else if (targetPage === "admin") fileName = "form-view";
  else if (targetPage === "qtv") fileName = "admin";
  else return '<div class="error">Trang không tồn tại!</div>';

  const template = HtmlService.createTemplateFromFile(fileName);
  template.user = user;
  return template.evaluate().getContent();
}

function testNhanVien() {
  const nv = getNhanVienByTen("Nguyễn Mạnh Hoàng");
  Logger.log(JSON.stringify(nv));
}

// Cập nhật thêm doGet để nhận diện ?page=view (nếu bạn chưa có)
function doGet(e) {
  const page =
    e && e.parameter && e.parameter.page ? e.parameter.page : "login";

  const routes = {
    login: { file: "login", title: "Trang đăng nhập CNTT" },
    admin: { file: "form-view", title: "Quản lý Ticket CNTT - Nhân viên IT" },
    qtv: {
      file: "admin",
      title: "Quản lý hệ thống Ticket CNTT - Admin quản trị viên",
    },
    view: { file: "form-view", title: "Danh sách Ticket của tôi" },
    create: { file: "form-create", title: "Trang gửi Ticket CNTT" },
  };

  const route = routes[page] || routes["login"];

  const htmlOutput = HtmlService.createHtmlOutputFromFile(route.file)
    .setTitle(route.title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  // Optional: Thêm debug (xem log khi load blank)
  Logger.log("doGet called with page: " + page + " | file: " + route.file);
  return htmlOutput;
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getLoaiSuCo() {
  const allOptions = { ...OPTIONS.LOAI_SU_CO }; // Lấy danh sách mặc định
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      // Bỏ qua dòng tiêu đề, lấy từ dòng 2
      for (let i = 1; i < data.length; i++) {
        const [ma, ten, bg, fg] = data[i];
        if (ma) {
          allOptions[ma] = { bg: bg, fg: fg, label: ten };
        }
      }
    }
  } catch (e) {
    console.log("Lỗi lấy danh mục: " + e.message);
  }
  return allOptions;
}

function processForm(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sheet = getSheet();
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    // 1. Xử lý File đính kèm
    const attachments = [];
    if (data.files && data.files.length > 0) {
      for (let file of data.files) {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(file.data),
          file.type,
          file.name,
        );
        const driveFile = folder.createFile(blob);
        driveFile.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW,
        );
        const id = driveFile.getId();
        const directUrl = `https://drive.google.com/uc?export=view&id=${id}`;

        if (file.type.startsWith("image/")) {
          attachments.push({
            formula: `=IMAGE("${directUrl}"; 1)`,
            link: driveFile.getUrl(),
          });
        } else {
          attachments.push({
            formula: driveFile.getUrl(),
            link: driveFile.getUrl(),
          });
        }
      }
    }

    // 2. TẠO STT VÀ MÃ SỰ CỐ
    const lastRow = sheet.getLastRow();
    let stt = 1;
    if (lastRow > 1) {
      const lastSttValue = sheet.getRange(lastRow, 1).getValue();
      stt = typeof lastSttValue === "number" ? lastSttValue + 1 : lastRow;
    }
    const sttFormatted = String(stt).padStart(4, "0");

    const ngayChuoi = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyyMMdd",
    );

    // ====== FIX PREFIX ======
    let prefix = data.loaiSuCo || "INC";
    prefix = prefix.toUpperCase().trim();

    const allowedPrefix = ["HIS", "BIS", "RIS", "LIS", "PHIS"];
    if (!allowedPrefix.includes(prefix)) prefix = "INC";
    const maSuCo = `${prefix}-${ngayChuoi}-${sttFormatted}`;

    // === MỚI: Tạo mã danh mục theo định dạng loai-ddmmyyyy ===
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, "0");
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy = today.getFullYear();
    const ddmmyyyy = dd + mm + yyyy;

    // Lấy mã danh mục gốc mà user chọn (từ frontend gửi lên)
    let maGoc = data.maDanhMucSuCo || "";

    // Nếu user chọn OTHER thì tạo danh mục mới
    if (maGoc === "OTHER") {
      const newMa = saveDanhMucIfOther(prefix, maGoc, data.moTaSuCo);

      if (newMa) {
        maGoc = newMa; // dùng mã mới
      }
    }

    // Parse mã gốc để tách loại và stt
    let loai = prefix;
    let sttGoc = "00";

    if (maGoc.includes("-")) {
      const parts = maGoc.split("-");
      loai = parts[0];
      sttGoc = parts[1];
    }
    const maDanhMucSuCo = `${loai}-${ddmmyyyy}-${sttGoc}`;

    // Set up ngày giờ mặc định theo hệ thống
    let ngayPhatSinhDate = new Date();
    if (data.ngayPhatSinh) {
      const parsed = new Date(data.ngayPhatSinh);
      if (!isNaN(parsed.getTime())) {
        ngayPhatSinhDate = parsed;
      }
    }

    // 3. CHUẨN BỊ DÒNG DỮ LIỆU MỚI
    const newRow = [
      stt, // A
      maSuCo, // B
      ngayPhatSinhDate, // C
      data.nguoiBaoSuCo || "", // D
      data.soDienThoai ? "'" + data.soDienThoai.trim() : "", // E
      data.boPhan || "", // F
      data.loaiSuCo || "", // G
      DEFAULT_STATUS.MUC_DO, // H
      data.maBenhNhan || "", // I
      data.hoTenBenhNhan || "", // J
      data.ngayChiDinh ? new Date(data.ngayChiDinh) : "", // K
      data.tenChiDinh || "", // L
      maDanhMucSuCo, // M - MÃ DANH MỤC NHẬN DẠNG MỚI: HIS-ddmmyyyy
      data.moTaSuCo || "", // N
      "", // 0 (File IT)
      DEFAULT_STATUS.NGUOI_TIEP_NHAN, // P
      "", // Q Nguyên nhân
      "", // R Hướng xử lý
      "", // S Thời gian bắt đầu
      "", // T Thời gian kết thúc
      "", // U Phút xử lý
      DEFAULT_STATUS.TRANG_THAI, // V
    ];

    Logger.log("Số cột newRow = " + newRow.length);

    sheet.appendRow(newRow);
    const rowIdx = sheet.getLastRow();

    sheet.autoResizeColumns(1, sheet.getLastColumn());
    sheet.getDataRange().setWrap(true);

    // 4. ĐỊNH DẠNG VÀ VALIDATION
    sheet.getRange(rowIdx, 3).setNumberFormat("dd/mm/yyyy HH:mm"); // Cột C
    sheet.getRange(rowIdx, 11).setNumberFormat("dd/mm/yyyy HH:mm"); // Cột K

    // Dropdowns
    setValidationAndStyle(
      sheet.getRange(rowIdx, 7),
      OPTIONS.LOAI_SU_CO,
      data.loaiSuCo,
    );

    // ✅ FIX: Xử lý file đính kèm - Hiển thị ảnh trực tiếp
    if (attachments.length > 0) {
      const cell = sheet.getRange(rowIdx, 23); // Cột W

      // Sử dụng setFormula thay vì setValue
      const fileInfo = attachments[0];
      cell.setFormula(fileInfo.formula);
      // Lưu link vào NOTE (cực kỳ quan trọng)
      cell.setNote(fileInfo.link);

      // Tăng chiều cao dòng để hiển thị ảnh rõ
      sheet.setRowHeight(rowIdx, 100);

      // Set alignment để ảnh đẹp hơn
      cell.setVerticalAlignment("middle");
      cell.setHorizontalAlignment("center");
    }

    // Định dạng thời gian
    sheet.getRange(rowIdx, 19, 1, 2).setNumberFormat("dd/mm/yyyy HH:mm"); // Cột S, T

    return { success: true, incidentId: maSuCo };
  } catch (err) {
    console.error(err);
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Hàm hỗ trợ tạo Dropdown
 */
function setValidationAndStyle(cell, optionConfig, selectedValue) {
  const options = Object.keys(optionConfig);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .build();
  cell.setDataValidation(rule);

  // Áp dụng màu sắc nếu giá trị tồn tại trong cấu hình
  const style = optionConfig[selectedValue];
  if (style) {
    cell.setBackground(style.bg);
    cell.setFontColor(style.fg);
    cell.setFontWeight("bold");
    cell.setHorizontalAlignment("left");
  }
}

function setDropdownWithStyle(cell, optionObj, value) {
  // Danh sách dropdown
  const list = Object.keys(optionObj);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  cell.setDataValidation(rule);
  cell.setValue(value || "");
  // Reset màu trước
  cell.setBackground(null);
  cell.setFontColor(null);
  // Nếu có màu trong OPTIONS → tô màu
  if (value && optionObj[value]) {
    cell.setBackground(optionObj[value].bg).setFontColor(optionObj[value].fg);
  }
}

function addLoaiSuCo(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    // Ghi dữ liệu mới vào dòng cuối
    sheet.appendRow([
      data.ma,
      data.ten,
      data.bg || "#e0e7ff",
      data.fg || "#3730a3",
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getTickets() {
  try {
    const sheet = getSheet();
    const range = sheet.getDataRange();

    // Lấy displayValues để giữ đúng format (quan trọng cho SĐT)
    const values = range.getDisplayValues();

    if (values.length <= 1) {
      return { success: true, tickets: [] };
    }

    const tickets = [];

    for (let i = 1; i < values.length; i++) {
      const r = values[i];

      // Bỏ qua dòng không có mã sự cố
      if (!r[1] || r[1].toString().trim() === "") continue;

      /* =====================
         LẤY FILE (CỘT W = 23)
      ====================== */
      let fileValue = "";

      try {
        const fileCell = sheet.getRange(i + 1, 23); // CỘT W

        const note = fileCell.getNote();
        if (note && note.trim() !== "") {
          fileValue = note.trim();
        } else {
          const formula = fileCell.getFormula();
          if (formula) {
            const match = formula.match(/https:\/\/drive\.google\.com[^\s")]+/);
            if (match) fileValue = match[0];
          }
        }
      } catch (e) {
        Logger.log("Lỗi file row " + (i + 1) + ": " + e);
      }

      /* =====================
         XỬ LÝ SỐ ĐIỆN THOẠI
      ====================== */
      let phone = (r[4] || "").toString().trim();

      // Nếu Google tự convert scientific notation (ví dụ 9.12345E+8)
      if (phone.includes("E+")) {
        phone = "'" + phone;
      }

      /* =====================
         PUSH DATA
      ====================== */

      tickets.push({
        stt: r[0],
        maSuCo: r[1],
        ngayPhatSinh: r[2],
        nguoiBaoSuCo: r[3],
        soDienThoai: phone,
        maKhoaPhong: r[5],
        loaiSuCo: r[6],
        mucDoUuTien: r[7],
        maBenhNhan: r[8],
        tenBenhNhan: r[9],
        ngayChiDinh: r[10],
        tenChiDinh: r[11],
        maDanhMucSuCo: r[12] || "", // mã danh mục sự cố
        moTaSuCo: r[13],
        nguoiTiepNhan: r[15],
        nguyenNhan: r[16],
        huongXuLy: r[17],
        thoiGianBatDau: r[18],
        thoiGianHoanTat: r[19],
        thoiGianXuLy: r[20],
        trangThai: r[21],
        file: fileValue,
        ghiChu: r[23],
      });
    }
    // Thêm map danh mục
    const danhMucMap = getDanhMucSuCoMap();

    return {
      success: true,
      tickets: tickets.reverse(),
    };
  } catch (err) {
    Logger.log("❌ Lỗi getTickets: " + err);
    return {
      success: false,
      message: err.message,
      tickets: [],
    };
  }
}

function uploadFile(base64, fileName, mimeType) {
  try {
    const folderId = "1otAxAX3cPs4lVr4W_FqSslI6SEf2XS7b";
    const folder = DriveApp.getFolderById(folderId);

    // Nếu base64 có dạng data:image/...
    if (base64.includes(",")) {
      base64 = base64.split(",")[1];
    }

    // Decode
    const bytes = Utilities.base64Decode(base64);

    const blob = Utilities.newBlob(
      bytes,
      mimeType || "application/octet-stream",
      fileName,
    );

    // Tạo file
    const file = folder.createFile(blob);

    // Public
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();

    // Link hiển thị ảnh
    const url = `https://drive.google.com/uc?export=view&id=${fileId}`;

    return {
      success: true,
      url: url,
      fileId: fileId,
    };
  } catch (e) {
    console.error("Upload error:", e);
    return {
      success: false,
      message: e.message,
    };
  }
}

function updateTicket(data) {
  try {
    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();

    // ✅ Biến để track xem có tìm thấy ticket không
    let found = false;

    for (let i = 1; i < values.length; i++) {
      // So khớp mã sự cố (Cột B = index 1)
      if (values[i][1] === data.maSuCo) {
        found = true; // Đánh dấu đã tìm thấy
        const row = i + 1;
        /* ================= NGƯỜI TIẾP NHẬN (P - 16) ================= */
        if (data.nguoiTiepNhan !== undefined) {
          setDropdownWithStyle(
            sheet.getRange(row, 16),
            OPTIONS.NGUOI_TIEP_NHAN,
            data.nguoiTiepNhan,
          );
        }

        /* ================= NGUYÊN NHÂN (Q - 17) ================= */
        if (data.nguyenNhan !== undefined) {
          sheet.getRange(row, 17).setValue(data.nguyenNhan || "");
        }

        /* ================= HƯỚNG XỬ LÝ (R - 18) ================= */
        if (data.huongXuLy !== undefined) {
          sheet.getRange(row, 18).setValue(data.huongXuLy || "");
        }

        /* ================= BẮT ĐẦU (S - 19) ================= */
        if (data.thoiGianBatDau) {
          sheet
            .getRange(row, 19)
            .setValue(parseVNDateTime(data.thoiGianBatDau));
        }

        /* ================= HOÀN TẤT (T - 20) ================= */
        if (data.thoiGianHoanTat) {
          sheet
            .getRange(row, 20)
            .setValue(parseVNDateTime(data.thoiGianHoanTat));
        }

        /* ================= TÍNH PHÚT (U - 21) ================= */
        if (data.thoiGianBatDau && data.thoiGianHoanTat) {
          const diff = calculateDiffMinutes(
            data.thoiGianBatDau,
            data.thoiGianHoanTat,
          );
          sheet.getRange(row, 21).setValue(diff);
        }

        /* ================= TRẠNG THÁI (V - 22) ================= */
        if (data.trangThai !== undefined) {
          setDropdownWithStyle(
            sheet.getRange(row, 22),
            OPTIONS.TRANG_THAI,
            data.trangThai,
          );
        }

        /* ================= MỨC ĐỘ ƯU TIÊN (H - 8) ================= */
        if (data.mucDoUuTien !== undefined) {
          setDropdownWithStyle(
            sheet.getRange(row, 8),
            OPTIONS.MUC_DO,
            data.mucDoUuTien,
          );
        }

        /* ================= FILE ĐÍNH KÈM (W - 23) ================= */
        if (data.fileBase64 && data.fileName) {
          try {
            const folderId = "1otAxAX3cPs4lVr4W_FqSslI6SEf2XS7b";
            const folder = DriveApp.getFolderById(folderId);

            let base64Clean = data.fileBase64;
            if (base64Clean.includes(",")) {
              base64Clean = base64Clean.split(",")[1];
            }

            const bytes = Utilities.base64Decode(base64Clean);
            const blob = Utilities.newBlob(
              bytes,
              data.mimeType || "application/octet-stream",
              data.fileName,
            );

            const file = folder.createFile(blob);

            file.setSharing(
              DriveApp.Access.ANYONE_WITH_LINK,
              DriveApp.Permission.VIEW,
            );

            const fileId = file.getId();
            const directUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
            const fileUrl = file.getUrl();

            const cell = sheet.getRange(row, 23); // ✅ CỘT W

            if (data.mimeType && data.mimeType.startsWith("image/")) {
              cell.setFormula(`=IMAGE("${directUrl}"; 1)`);
              cell.setNote(fileUrl);
            } else {
              cell.setValue(directUrl);
            }

            sheet.setRowHeight(row, 80);
          } catch (fileErr) {
            throw new Error("Không thể upload file: " + fileErr.message);
          }
        }

        /* ================= GHI CHÚ (X - 24) ================= */
        if (data.ghiChu !== undefined) {
          sheet.getRange(row, 24).setValue(data.ghiChu || "");
        }
        break;
      }
    }
    if (found) {
      return {
        success: true,
        message: `✅ Cập nhật ticket ${data.maSuCo} thành công!`,
      };
    } else {
      return {
        success: false,
        message: "❌ Không tìm thấy mã sự cố!",
      };
    }
  } catch (err) {
    console.error("updateTicket error:", err);
    return {
      success: false,
      message: "❌ Lỗi hệ thống: " + err.message,
    };
  }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("❌ Không tìm thấy sheet: " + SHEET_NAME);
  }
  return sheet;
}

function getSheetUser() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USER);
  if (!sheet) {
    throw new Error("❌ Không tìm thấy sheet: ");
  } else {
    console.log("Tìm thấy sheet users");
  }
  return sheet;
}

function getAllHoTen() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NS); // đổi đúng tên sheet

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const name = data[i][6]; // Cột G

    if (name && name.toString().trim() !== "") {
      list.push(name.toString().trim());
    }
  }
  return list;
}

function getAllTrangThai() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("TrangThai");
    if (!sheet) return ["Chưa xử lý", "Đang xử lý", "Hoàn tất"]; // fallback nếu sheet không tồn tại

    const data = sheet.getDataRange().getValues();
    const result = [];

    for (let i = 1; i < data.length; i++) {
      // bỏ header
      let tenTrangThai = (data[i][1] || "").toString().trim(); // Cột B
      // Loại bỏ số thứ tự ở đầu (ví dụ: "1 Chưa xử lý" → "Chưa xử lý")
      tenTrangThai = tenTrangThai.replace(/^\d+\s*/, "").trim();

      if (tenTrangThai) {
        result.push(tenTrangThai);
      }
    }

    return result;
  } catch (e) {
    console.error("Lỗi getAllTrangThai:", e);
    return ["Chưa xử lý", "Đang xử lý", "Hoàn tất"];
  }
}

function getNhanVienByMa(maNV) {
  if (!maNV) return null;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const ma = String(data[i][1]).trim(); // CỘT B = MANV
    const maBP = String(data[i][3]).trim(); // CỘT C = MABP
    const hoTen = String(data[i][6]).trim(); // CỘT G = HOTEN
    if (ma && ma.toString().trim() === maNV) {
      return {
        maNV: ma,
        hoTen: hoTen || "",
        maBoPhan: maBP || "",
      };
    }
  }
  return null;
}

function test_getNhanVienByTen() {
  const result = getNhanVienByTen("Nguyễn Mạnh Hoàng");
  Logger.log(result);
}

function normalizeString(str) {
  return str.toString().normalize("NFC").trim().toLowerCase();
}

function getNhanVienByTen(ten) {
  if (!ten) return null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NS);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const inputTen = normalizeString(ten);

  for (let i = 1; i < data.length; i++) {
    const ma = String(data[i][1]).trim(); // B
    const maBP = String(data[i][2]).trim(); // C
    const hoTen = data[i][6];
    const soDienThoai = data[i][25]; // Z (cột 25)
    if (!hoTen) continue;
    const sheetTen = normalizeString(hoTen);
    if (sheetTen === inputTen) {
      return {
        maNV: ma,
        hoTen: hoTen.toString().trim(),
        maKhoaPhong: maBP || "",
        soDienThoai: soDienThoai ? soDienThoai.toString().trim() : "",
      };
    }
  }
  return null;
}

function getKhoaByMaNhanVien() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === maNV) {
      // cột A = MaNV
      return {
        maKhoaPhong: data[i][2], // cột C
        tenKhoa: data[i][3], // cột D
      };
    }
  }
  return null;
}

function getAllBoPhan() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const set = new Set(); // 👈 Dùng Set để tự loại trùng
  for (let i = 1; i < data.length; i++) {
    const name = data[i][3]; // Cột D
    if (name && name.toString().trim() !== "") {
      set.add(name.toString().trim());
    }
  }
  // Chuyển Set → Array
  return Array.from(set);
}

function getAllMaBoPhan() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const set = new Set();
  for (let i = 1; i < data.length; i++) {
    const ma = data[i][2]; // Cột C
    if (ma && ma.toString().trim() !== "") {
      set.add(ma.toString().padStart(3, "0"));
    }
  }
  return Array.from(set);
}

function testBoPhan() {
  Logger.log(getAllMaBoPhan());
}

// Lấy map danh mục sự cố
function getDanhMucSuCoMap() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("DanhMucSuCo");
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < data.length; i++) {
    // Bỏ header
    const ma = (data[i][0] || "").toString().trim(); // Cột A: HIS-01, BIS-01,...
    const moTa = (data[i][1] || "").toString().trim(); // Cột B: mô tả
    if (ma && moTa) map[ma] = moTa;
  }
  return map;
}

// Lấy danh sách mã loại sự cố
function getMaSuCo(maLoai) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("DanhMucSuCo");
    if (!sheet) throw new Error("Không tìm thấy sheet DanhMucSuCo");

    const data = sheet.getDataRange().getValues();
    const result = [];

    // Bắt đầu từ hàng 2 (bỏ header)
    for (let i = 1; i < data.length; i++) {
      const maChiTiet = (data[i][0] || "").toString().trim(); // Cột A mới: HIS-01, BIS-01,...
      const tenDanhMuc = (data[i][1] || "").toString().trim(); // Cột B: tên danh mục

      if (maChiTiet.startsWith(maLoai + "-") && tenDanhMuc) {
        result.push({
          ma: maChiTiet, // HIS-01
          ten: tenDanhMuc, // Trùng số thứ tự chờ khám
        });
      }
    }

    // Sắp xếp theo mã (tùy chọn, để đẹp hơn)
    result.sort((a, b) => a.ma.localeCompare(b.ma));
    return result;
  } catch (e) {
    Logger.log("Lỗi: " + e);
    return [];
  }
}

// Lấy danh sách phòng khoa ứng với mã
function getKhoaPhongMap() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("BoPhan"); // Tạo sheet này nếu chưa có
  if (!sheet) return {}; // Nếu sheet chưa tồn tại, trả về rỗng

  const data = sheet.getDataRange().getValues();
  const map = {};

  // Bỏ header (dòng 1)
  for (let i = 1; i < data.length; i++) {
    const ma = data[i][0] ? String(data[i][0]).trim() : null; // mã khoa/phòng
    const ten = data[i][1] ? data[i][1].trim() : null; // tên khoa/phòng
    if (ma && ten) {
      map[ma] = ten;
    }
  }
  return map;
}

function getAllLoaiSuCo() {
  try {
    // Nếu hàm này nằm trong cùng file Spreadsheet thì có thể dùng SpreadsheetApp.getActive()
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("LoaiSuCo");
    if (!sheet) {
      throw new Error("Không tìm thấy sheet 'LoaiSuCo'");
    }

    const data = sheet.getDataRange().getValues();
    const result = [];

    // Bắt đầu từ hàng 2 (bỏ header)
    for (let i = 1; i < data.length; i++) {
      const ma = (data[i][1] || "").toString().trim(); // Cột B: mã (HIS, BIS,...)
      const ten = (data[i][2] || "").toString().trim(); // Cột C: tên đầy đủ
      if (ma && ten) {
        result.push({
          ma: ma,
          ten: ten,
        });
      }
    }

    // Sắp xếp theo mã (tùy chọn)
    result.sort((a, b) => a.ma.localeCompare(b.ma));
    return result;
  } catch (e) {
    Logger.log("Lỗi getAllLoaiSuCo: " + e.message);
    return [];
  }
}

function testGetTickets() {
  const result = getTickets();
  Logger.log("Success: " + result.success);
  Logger.log("Total tickets: " + result.tickets.length);

  if (result.tickets.length > 0) {
    const firstTicket = result.tickets[0];
    Logger.log("First ticket file: " + firstTicket.file);
  }
}

// Lấy danh sách danh mục sự cố theo loại (mã sự cố)
function getDanhMucByLoai(maLoai) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheetMap = {
    HIS: "DanhMuc_HIS",
    BIS: "DanhMuc_BIS",
    RIS: "DanhMuc_RIS",
    LIS: "DanhMuc_LIS",
  };

  const sheetName = sheetMap[(maLoai || "").toUpperCase()];
  if (!sheetName) return [];

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return data.map((r) => ({
    ma: r[0],
    ten: r[1],
  }));
}

// Cache danh mục sự cố (dùng CacheService, hết hạn sau 10 phút)
function getDanhMucByLoaiCached(maLoai) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "danhMuc_" + maLoai.toUpperCase();
  let cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const result = getDanhMucByLoai(maLoai);
  cache.put(cacheKey, JSON.stringify(result), 600); // cache 10 phút
  return result;
}

// Clear (xóa) cache
// function clearDanhMucCache(loai) {
//   const cache = CacheService.getScriptCache();
//   const key = "danhmuc_" + loai;
//   cache.remove(key);
// }

// Hàm gửi zalo
// function sendZaloNotify(message) {
//   // Lấy chuỗi accessToken từ zalo
//   const accessToken = "YOUR_ZALO_ACCESS_TOKEN";
//   const payload = {
//     recipient: {
//       user_id: "USER_ID_IT"
//     },
//     message: {
//       text: message
//     }
//   };

//   const options = {
//     method: "post",
//     contentType: "application/json",
//     headers: {
//       access_token: accessToken
//     },
//     payload: JSON.stringify(payload)
//   };

//   const url = "https://openapi.zalo.me/v3.0/oa/message/cs";

//   UrlFetchApp.fetch(url, options);

// }

// Lưu danh mục vào sheet danh mục cụ thể khi chọn OTHER => nhập danh mục cụ thể
function saveDanhMucIfOther(loaiSuCo, maDanhMucSuCo, moTaSuCo) {
  // Nếu không phải chọn mục OTHER thì bỏ qua
  if (maDanhMucSuCo !== "OTHER") return null;
  if (!loaiSuCo || !moTaSuCo) return null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = "DanhMuc_" + loaiSuCo;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["MaDanhMuc", "MoTa"]);
  }
  const data = sheet.getDataRange().getValues();
  const moTaInput = moTaSuCo.trim().toLowerCase();

  // 🔎 kiểm tra trùng mô tả danh mục sự cố thường gặp
  for (let i = 1; i < data.length; i++) {
    const moTaExist = data[i][1].toString().trim().toLowerCase();
    if (moTaExist === moTaInput) {
      // nếu mô tả danh mục sự cố bị trùng → dùng lại mã cũ
      return data[i][0];
    }
  }

  // không trùng → tạo mã danh mục mới
  const nextNumber = sheet.getLastRow(); // vì có header
  const maDanhMuc = loaiSuCo + "-" + String(nextNumber).padStart(2, "0");

  sheet.appendRow([maDanhMuc, moTaSuCo.trim()]);

  return maDanhMuc;
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const result = processForm(data);
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
