function clasificarFeedback() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Respuestas de formulario 1",
  );
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];

  const colComentario = encabezados.indexOf("Comentario/Feedback");
  const colSentimiento = encabezados.indexOf("Categoria de Sentimiento");
  const colResumen = encabezados.indexOf("Resumen IA");

  if (colComentario === -1 || colSentimiento === -1 || colResumen === -1) {
    Logger.log(
      "ERROR: No se encontró alguna columna. Revisa los nombres exactos de los encabezados.",
    );
    Logger.log("Encabezados encontrados: " + JSON.stringify(encabezados));
    return;
  }

  const apiKey =
    PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const comentario = fila[colComentario];
    const sentimientoActual = fila[colSentimiento];

    if (sentimientoActual !== "" && sentimientoActual !== undefined) {
      continue;
    }

    if (!comentario) {
      continue;
    }

    const resultado = clasificarConGemini(comentario, apiKey);

    if (resultado) {
      hoja.getRange(i + 1, colSentimiento + 1).setValue(resultado.sentimiento);
      hoja.getRange(i + 1, colResumen + 1).setValue(resultado.resumen);
    }

    Utilities.sleep(1000);
  }

  Logger.log("Proceso terminado.");
}

function clasificarConGemini(comentario, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;

  const prompt = `Analiza el siguiente comentario de feedback de un cliente y responde SOLO con un JSON válido, sin texto adicional, sin markdown, con este formato exacto:
{"sentimiento": "Positivo" o "Neutro" o "Negativo", "resumen": "resumen breve de una frase"}






Comentario: "${comentario}"`;

  const payload = {
    contents: [
      {
        parts: [{ text: prompt }],
      },
    ],
  };

  const opciones = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const respuesta = UrlFetchApp.fetch(url, opciones);
    const json = JSON.parse(respuesta.getContentText());

    Logger.log("Respuesta completa de Gemini: " + JSON.stringify(json));

    let textoIA = json.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g, "").trim();

    const resultado = JSON.parse(textoIA);
    return resultado;
  } catch (error) {
    Logger.log("Error clasificando comentario: " + error);
    return null;
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Formulario")
    .setTitle("Feedback Alegra")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function guardarFeedback(datos) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Respuestas de formulario 1",
  );
  const encabezados = hoja
    .getRange(1, 1, 1, hoja.getLastColumn())
    .getValues()[0];

  const colComentario = encabezados.indexOf("Comentario/Feedback");
  const colProducto = encabezados.indexOf("Producto");
  const colNombre = encabezados.indexOf("Nombre del Usuario");
  const colSentimiento = encabezados.indexOf("Categoria de Sentimiento");
  const colResumen = encabezados.indexOf("Resumen IA");

  const nuevaFila = new Array(encabezados.length).fill("");

  nuevaFila[0] = new Date();
  if (colNombre !== -1) nuevaFila[colNombre] = datos.nombre;
  if (colProducto !== -1) nuevaFila[colProducto] = datos.producto;
  if (colComentario !== -1) nuevaFila[colComentario] = datos.comentario;

  hoja.appendRow(nuevaFila);

  // Clasificamos inmediatamente el comentario recién guardado
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  Logger.log("API Key encontrada: " + (apiKey ? "SÍ" : "NO, está vacía"));

  const resultado = clasificarConGemini(datos.comentario, apiKey);
  Logger.log("Resultado de clasificación: " + JSON.stringify(resultado));

  if (resultado) {
    const numeroFila = hoja.getLastRow();
    if (colSentimiento !== -1)
      hoja
        .getRange(numeroFila, colSentimiento + 1)
        .setValue(resultado.sentimiento);
    if (colResumen !== -1)
      hoja.getRange(numeroFila, colResumen + 1).setValue(resultado.resumen);
    Logger.log("Columnas escritas correctamente en la fila " + numeroFila);
  } else {
    Logger.log("El resultado fue null, no se escribió nada.");
  }

  return "ok";
}
