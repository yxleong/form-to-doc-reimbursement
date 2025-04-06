function onFormSubmit(e) {
  const templateId = 'REPLACE_DOC_ID';
  const folderId = 'REPLACE_NEW_DOC_FOLDER_ID';
  const responses = e.values;
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  let data = {};
  for (let i = 0; i < headers.length; i++) {
    data[headers[i]] = responses[i];
  }

// to meet specific requirements ---------------------------------------------------------------------------------------------
  data["日期"] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  const codeParts = (data["預算科目代號"] || "").split("-");
  data["K"] = codeParts[0] || "";
  data["X"] = codeParts[1] || "";
  data["M"] = codeParts[2] || "";
  data["J"] = codeParts[3] || "";

  const amount = data["金額"] || "";
  const amtChars = amount.split("");
  data["HT"] = amtChars[0] || "";
  data["TT"] = amtChars[1] || "";
  data["T"] = amtChars[2] || "";
  data["H"] = amtChars[3] || "";
  data["TEN"] = amtChars[4] || "";
  data["Z"] = amtChars[5] || "";

  const now = new Date();
  const nowtimestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMddHHmmss");
  const fileName = `${nowtimestamp}核銷單-${data["承辦人"]}`;
// ---------------------------------------------------------------------------------------------------------------------------

  const copy = DriveApp.getFileById(templateId).makeCopy(fileName, folderId ? DriveApp.getFolderById(folderId) : undefined);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  for (const [key, value] of Object.entries(data)) {
    body.replaceText(`{{${key}}}`, value);
  }

// to meet specific requirements ---------------------------------------------------------------------------------------------
  try {
    const urlString = data["發票或收據"];
    const urls = urlString.split(/\s*,\s*|\n/).filter(Boolean);
    const imageBlobs = [];

    for (let url of urls) {
      const fileIdMatch = url.match(/[-\w]{25,}/);
      if (fileIdMatch) {
        const fileId = fileIdMatch[0];
        const blob = DriveApp.getFileById(fileId).getBlob();
        imageBlobs.push(blob);
      }
    }

    body.replaceText('{{QTY}}', imageBlobs.length.toString());

    const maxWidth = 15 * 37.795;
    const maxHeight = 25 * 37.795;

    if (imageBlobs.length > 0) {
      let foundElement;
      while ((foundElement = body.findText('{{image}}'))) {
        const parent = foundElement.getElement().getParent();
        parent.clear();

        imageBlobs.forEach(blob => {
          const image = parent.appendInlineImage(blob);
          const width = image.getWidth();
          const height = image.getHeight();

          const scale = Math.min(maxWidth / width, maxHeight / height);
          image.setWidth(width * scale);
          image.setHeight(height * scale);
        });
      }
    }
  } catch (err) {
    body.replaceText('{{QTY}}', '0');
  }
// ---------------------------------------------------------------------------------------------------------------------------

  doc.saveAndClose();

  const pdf = copy.getAs(MimeType.PDF);

  MailApp.sendEmail({
    to: data["電子郵件地址"],
    subject: "國立臺灣科技大學學生會--原始憑證黏存單",
    body:  `敬啟者您好，

感謝您提交本次核銷申請，相關資料已成功建立！
核銷表單PDF副本已隨函附上，請您查閱確認。

若對內容有任何疑問或需補充資料，請於三日內回覆本郵件。
如無其他問題，我們將於七日內完成核銷處理，屆時會通知您後續進度。

敬祝
順心如意

──────────────────────────────
Best regards,
國立臺灣科技大學 學生會 財務部
Student Association, Department of Finance
National Taiwan University of Science and Technology
──────────────────────────────
  `,
    attachments: [pdf]
  });
}
