const SHEET_NAME = "Trang tính1";
const DRIVE_FOLDER_ID = "1otAxAX3cPs4lVr4W_FqSslI6SEf2XS7b";

// --- CẤU HÌNH DANH SÁCH DROPDOWN VÀ MÀU SẮC ---
const OPTIONS = {
  LOAI_SU_CO: {
    HIS: { bg: "#d9ead3", fg: "#274e13" },
    BIS: { bg: "#cfe2f3", fg: "#073763" },
    RIS: { bg: "#d0e0e3", fg: "#134f5c" },
    LIS: { bg: "#c9daf8", fg: "#1c4587" },
    PHIS: { bg: "#ead1dc", fg: "#741b47" },
  },
  MUC_DO: {
    "P1 - Khẩn cấp": { bg: "#f4cccc", fg: "#990000" },
    "P2 - Cao": { bg: "#fff2cc", fg: "#7f6000" },
    "P3 - Thường": { bg: "#d9ead3", fg: "#274e13" },
  },
  NGUOI_TIEP_NHAN: {
    Tuấn: { bg: "#cfe2f3", fg: "#073763" },
    Hùng: { bg: "#d9ead3", fg: "#274e13" },
    Khoa: { bg: "#ead1dc", fg: "#741b47" },
    Huy: { bg: "#fff2cc", fg: "#7f6000" },
  },
};

function doGet(e) {
  const page =
    e && e.parameter && e.parameter.page ? e.parameter.page : "create";
  if (page === "admin") {
    return HtmlService.createHtmlOutputFromFile("form-view").setTitle(
      "Quản lý Ticket CNTT",
    );
  }
  return HtmlService.createHtmlOutputFromFile("form-create").setTitle(
    "Gửi Ticket CNTT",
  );
}

function processForm(data) {
  try {
    const sheet = getSheet();
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    // 1. Xử lý File đính kèm (giữ nguyên)
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
        const url = driveFile.getUrl();
        const id = driveFile.getId();
        if (file.type.startsWith("image/")) {
          attachments.push({
            formula: `=IMAGE("https://drive.google.com/uc?export=view&id=${id}";1)`,
            link: url,
          });
        } else {
          attachments.push({ formula: url, link: url });
        }
      }
    }

    // 2. FIX LỖI TẠO STT VÀ MÃ SỰ CỐ
    const lastRow = sheet.getLastRow();
    let stt = 1;

    if (lastRow > 1) {
      // Lấy giá trị STT ở cột A của dòng cuối cùng hiện có
      const lastSttValue = sheet.getRange(lastRow, 1).getValue();
      // Nếu là số thì cộng 1, nếu không thì dựa vào lastRow
      stt = typeof lastSttValue === "number" ? lastSttValue + 1 : lastRow;
    }

    const sttFormatted = String(stt).padStart(4, "0");
    const ngayChuoi = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyyMMdd",
    );
    const maSuCo = `INC-${ngayChuoi}-${sttFormatted}`;

    // 3. Chuẩn bị dòng dữ liệu mới (Cột A bây giờ dùng biến stt vừa tính)
    const newRow = [
      stt, // Cột A: STT (Đã fix)
      maSuCo, // Cột B: Mã sự cố
      new Date(data.ngayPhatSinh), // Cột C: Ngày phát sinh
      data.nguoiBaoSuCo,
      data.boPhan,
      data.loaiSuCo,
      data.mucDoUuTien,
      data.moTaSuCo,
      data.nguoiTiepNhan,
      data.nguyenNhan || "",
      data.huongXuLy || "",
      data.thoiGianBatDau ? new Date(data.thoiGianBatDau) : "",
      data.thoiGianHoanTat ? new Date(data.thoiGianHoanTat) : "",
      Number(data.thoiGianXuLy || 0),
      data.trangThai || "Mới",
      "",
      data.ghiChu || "",
    ];

    sheet.appendRow(newRow);
    const rowIdx = sheet.getLastRow();

    // Định dạng (Ngày phát sinh dd/MM/yyyy và Dropdown - Giữ nguyên như cũ)
    sheet.getRange(rowIdx, 3).setNumberFormat("dd/MM/yyyy");
    setValidationAndStyle(
      sheet.getRange(rowIdx, 6),
      OPTIONS.LOAI_SU_CO,
      data.loaiSuCo,
    );
    setValidationAndStyle(
      sheet.getRange(rowIdx, 7),
      OPTIONS.MUC_DO,
      data.mucDoUuTien,
    );
    setValidationAndStyle(
      sheet.getRange(rowIdx, 9),
      OPTIONS.NGUOI_TIEP_NHAN,
      data.nguoiTiepNhan,
    );

    // Xử lý file (giữ nguyên)
    if (attachments.length > 0) {
      const cell = sheet.getRange(rowIdx, 16);
      if (attachments[0].formula.startsWith("=")) {
        cell.setFormula(attachments[0].formula);
        sheet.setRowHeight(rowIdx, 80);
      } else {
        cell.setValue(attachments[0].link);
      }
      if (attachments.length > 1) {
        cell.setNote(
          attachments.map((a, i) => `File ${i + 1}: ${a.link}`).join("\n"),
        );
      }
    }

    sheet.getRange(rowIdx, 12, 1, 2).setNumberFormat("dd/MM/yyyy HH:mm");
    return { success: true, incidentId: maSuCo };
  } catch (err) {
    console.error(err);
    return { success: false, error: err.toString() };
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

function getTickets() {
  try {
    const sheet = getSheet();
    const values = sheet.getDataRange().getDisplayValues();

    if (values.length <= 1) return { success: true, tickets: [] };

    const tickets = [];
    // Bắt đầu từ i = 1 để bỏ qua tiêu đề
    for (let i = 1; i < values.length; i++) {
      const r = values[i];

      // --- Kiểm tra nếu cột Mã sự cố (r[1]) hoặc Người báo (r[3]) trống thì bỏ qua ---
      if (!r[1] || r[1].trim() === "") {
        continue;
      }

      tickets.push({
        stt: r[0],
        maSuCo: r[1],
        ngayPhatSinh: r[2],
        nguoiBaoSuCo: r[3],
        boPhan: r[4],
        loaiSuCo: r[5],
        mucDoUuTien: r[6],
        moTaSuCo: r[7],
        nguoiTiepNhan: r[8],
        nguyenNhan: r[9],
        huongXuLy: r[10],
        thoiGianBatDau: r[11],
        thoiGianHoanTat: r[12],
        thoiGianXuLy: r[13],
        trangThai: r[14],
        file: r[15],
        ghiChu: r[16],
      });
    }

    // Trả về danh sách đã lọc, đảo ngược để hiện cái mới nhất lên đầu
    return { success: true, tickets: tickets.reverse() };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateTicket(data) {
  try {
    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === data.maSuCo) {
        const row = i + 1;
        sheet.getRange(row, 10).setValue(data.nguyenNhan || "");
        sheet.getRange(row, 11).setValue(data.huongXuLy || "");
        sheet.getRange(row, 14).setValue(Number(data.thoiGianXuLy || 0));
        sheet.getRange(row, 15).setValue(data.trangThai);
        sheet.getRange(row, 17).setValue(data.ghiChu || "");
        return { success: true };
      }
    }
    return { success: false, message: "Không tìm thấy mã" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("❌ Không tìm thấy sheet: " + SHEET_NAME);
  return sheet;
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const result = processForm(data);
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
