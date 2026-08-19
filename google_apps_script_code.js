// =========================================================================
// J&J VENTURES — GOOGLE APPS SCRIPT WEBHOOK & CALENDAR AUTOMATION
// =========================================================================
// Pega este código en: Google Sheets -> Extensiones -> Apps Script -> Código.gs
// Luego clic en Implementar -> Gestionar implementaciones -> Editar -> Nueva versión -> Implementar.

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    // 1. Crear encabezados en la fila 1 si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Fecha Registro",
        "Nombre",
        "Email Corporativo",
        "Teléfono / WhatsApp",
        "Sitio Web / LinkedIn",
        "Fecha Cita",
        "Hora Cita",
        "Modelo Empresa",
        "Obstáculo Principal",
        "Ticket Promedio",
        "Facturación Mensual",
        "Tiempo Operando",
        "Urgencia",
        "Presupuesto"
      ]);
      sheet.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#0F172A").setFontColor("#FFFFFF");
      sheet.setFrozenRows(1);
    }
    
    // 2. Insertar fila con la data completa del prospecto
    sheet.appendRow([
      new Date(),
      data.fullName || "",
      data.workEmail || "",
      data.phone || "",
      data.website || "",
      data.meetingDateFormatted || data.meetingDate || "Por confirmar",
      data.meetingTime || "Por confirmar",
      data.businessType || "",
      data.mainObstacle || "",
      data.ticketAverage || "",
      data.monthlyRevenue || "",
      data.businessAge || "",
      data.urgency || "",
      data.budgetFit || ""
    ]);
    
    // 3. Crear Evento en Google Calendar e invitar al prospecto automáticamente
    if (data.meetingDate && data.meetingTime && data.workEmail) {
      try {
        var timeParts = data.meetingTime.match(/(\d+):(\d+)\s*(AM|PM)/i);
        if (timeParts) {
          var hours = parseInt(timeParts[1], 10);
          var minutes = parseInt(timeParts[2], 10);
          var ampm = timeParts[3].toUpperCase();
          if (ampm === "PM" && hours < 12) hours += 12;
          if (ampm === "AM" && hours === 12) hours = 0;
          
          var dateParts = data.meetingDate.split("-");
          var year = parseInt(dateParts[0], 10);
          var month = parseInt(dateParts[1], 10) - 1;
          var day = parseInt(dateParts[2], 10);
          
          var startDateTime = new Date(year, month, day, hours, minutes, 0);
          var endDateTime = new Date(startDateTime.getTime() + (15 * 60 * 1000)); // 15 min
          
          var cal = CalendarApp.getDefaultCalendar();
          var eventTitle = "Diagnóstico de Prospección B2B — J&J Ventures / " + data.fullName;
          
          var eventDescription = 
            "Sesión de Diagnóstico Outbound (15 min) — J&J Ventures\n\n" +
            "• Prospecto: " + data.fullName + "\n" +
            "• Email: " + data.workEmail + "\n" +
            "• Teléfono / WhatsApp: " + data.phone + "\n" +
            "• Modelo Empresa: " + data.businessType + "\n" +
            "• Obstáculo Comercial: " + data.mainObstacle + "\n" +
            "• Ticket Promedio: " + data.ticketAverage + "\n" +
            "• Facturación: " + data.monthlyRevenue + "\n\n" +
            "Nos conectaremos puntualmente por Google Meet / Zoom.\n" +
            "J&J Ventures · Prov. 16:3";
          
          cal.createEvent(eventTitle, startDateTime, endDateTime, {
            description: eventDescription,
            guests: data.workEmail,
            sendInvites: true
          });
        }
      } catch (calErr) {
        Logger.log("Error creando evento de calendario: " + calErr);
      }
    }
    
    // 4. Enviar Alerta Inmediata a tu propio Correo
    var operatorEmail = "joerneto26@gmail.com";
    MailApp.sendEmail({
      to: operatorEmail,
      subject: "🚨 NUEVA CITA AGENDADA — J&J Ventures: " + data.fullName + " (" + (data.meetingDateFormatted || data.meetingDate) + " " + data.meetingTime + ")",
      htmlBody: 
        "<h3>🎯 Nuevo prospecto ha agendado una llamada de diagnóstico:</h3>" +
        "<ul>" +
        "<li><b>Nombre:</b> " + data.fullName + "</li>" +
        "<li><b>Email:</b> " + data.workEmail + "</li>" +
        "<li><b>Teléfono:</b> " + data.phone + "</li>" +
        "<li><b>Fecha de Cita:</b> " + (data.meetingDateFormatted || data.meetingDate) + " a las " + data.meetingTime + "</li>" +
        "<li><b>Modelo Empresa:</b> " + data.businessType + "</li>" +
        "<li><b>Obstáculo:</b> " + data.mainObstacle + "</li>" +
        "<li><b>Ticket Promedio:</b> " + data.ticketAverage + "</li>" +
        "<li><b>Facturación:</b> " + data.monthlyRevenue + "</li>" +
        "<li><b>Presupuesto Confirmado:</b> " + data.budgetFit + "</li>" +
        "</ul>" +
        "<p><i>El evento ha sido registrado en tu Google Calendar y se envió la invitación al prospecto.</i></p>"
    });
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
