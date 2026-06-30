function makeQRCodes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Картриджи');
  const lastRow = sheet.getLastRow();
  const webAppUrl = "https://script.google.com/macros/s/*--*/exec";
 
  for (let i = 2; i <= lastRow; i++) {
    const id = sheet.getRange(i, 1).getValue().toString().trim();
    if (id) {
      const targetUrl = `${webAppUrl}?id=${id}`;
      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=75x75&data=${encodeURIComponent(targetUrl)}`;
      sheet.getRange(i, 9).setFormula(`=IMAGE("${qrUrl}")`);
    }
  }
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Картриджи");
  const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("История");
  const id = e.parameter.id;
  const action = e.parameter.action;
  const baseUrl = ScriptApp.getService().getUrl();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const row = i + 1;
      const model = data[i][1];      // Модель
      const printer = data[i][2];    // Принтер
      const dept = data[i][3];       // Кабинет / отдел
      //если просто просмотр карточки
      if (!action) {
        return HtmlService.createHtmlOutput(`
          <h2>Картридж ID: ${id}</h2>
          <p><b>Принтер:</b> ${printer}</p>
          <p><b>Отдел:</b> ${dept}</p>
          <hr>
          <a href="${baseUrl}?id=${id}&action=taken">
            <button>Взял на заправку</button>
          </a>
          <a href="${baseUrl}?id=${id}&action=returned">
            <button>Отдал в работу</button>
          </a>
        `);
      }
      //дата + время
      const now = Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm:ss"
      );
      let actionText = "";
      if (action === "taken") {
        sheet.getRange(row, 6).setValue(now);
        sheet.getRange(row, 5).setValue("В работе");
        actionText = "Взят на заправку";
      }
      if (action === "returned") {
        sheet.getRange(row, 7).setValue(now);
        sheet.getRange(row, 5).setValue("Готов");
        const current = Number(sheet.getRange(row, 8).getValue()) || 0;
        sheet.getRange(row, 8).setValue(current + 1);
        actionText = "Отдан в работу";
      }
      // история
      historySheet.appendRow([now, id, actionText]);

      // ответ
      return HtmlService.createHtmlOutput(`
        <h2>ОК</h2>
        <p>Сохранено</p>
        <script>
          setTimeout(() => {
            window.location.href = "${baseUrl}?id=${id}";
          }, 800);
        </script>
      `);
    }
  }
  return HtmlService.createHtmlOutput("Не найдено");
}
