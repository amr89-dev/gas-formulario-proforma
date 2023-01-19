const SS = SpreadsheetApp.openById(
  "1J1CEwIS9FMmsiNUsi-mxdzPLhGC4bgW4FCUMTwDcc34"
);

const ID_PLANTILLA = {
  1: "1VLd-OLLRhHJtQXvDaJdrBU_9cwy-_PN4aK1MvhbbkqU",
  2: "1wlGR5Kk1nN1P9SCtj6b7rw9UKBfwTMPVXcQ1mO31KdY",
  3: "1ldCqYCj4oa_qYO-hBju91RU0WopLc2fUjlgtUPEqN_o",
  4: "16yY-njFvzwlcim9ZRCAmUYXRi_VUHEwHDLswZU6kHWQ",
  5: "1oFj0GWTV5hibYfRqGqiyFUjC-NogAhzkuXe17bnI5_A",
  6: "1g92KnFktnq1-vsBTkTbVeKEw16Z7s_gkrhETr79t9wI",
  7: "1btOd6xdsWL3RNm9qZXC_pECNAB1nMJc93IislQ4QRHM",
  8: "10ae_Y8fKXvFr3yvvDuoCSt2U1buTMtC5qXchpE4eE4E",
  9: "1KD5WUVvUvqWS6kMOmrlDBBps3PIl2b6znFza5iNcMoE",
  10: "1VLd-OLLRhHJtQXvDaJdrBU_9cwy-_PN4aK1MvhbbkqU",
  11: "1K0CGN8AoJvSR8dtxojk79rGN7S84r2urprs6ltp0WNY",
  12: "1m8dtK4Js7os1DdJElxP40uZ8Su8E2fU1O_kUh9mZ4cg",
};

const sheetRegistro = SS.getSheetByName("registro");
const sheetPdf = SS.getSheetByName("pdf");
const sheetDataProductos = SS.getSheetByName("bd_lista");

function doGet() {
  var template = HtmlService.createTemplateFromFile("registro");
  template.pubUrl =
    "https://script.google.com/a/macros/losacordeones.com/s/AKfycbyLBFuQ3xIc4VtaYQyS1DvPxm4OADCwEAECKaay5FK1A2xPBE_jnTAzZX39YNdg6w99/exec";
  var output = template.evaluate();
  return output;
}

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}
function getDataProductos() {
  const dataProductos = sheetDataProductos.getDataRange().getDisplayValues();
  dataProductos.shift();
  console.log(dataProductos);
  return dataProductos;
}
function setCurrencyFormat() {
  const dateRange = sheetRegistro.getRangeList(["B2:B", "C2:C"]);
  dateRange.setNumberFormat("d/M/yyyy");

  const rangesToFormat = sheetRegistro.getRangeList([
    "N2:N",
    "O2:O",
    "S2:S",
    "T2:T",
    "X2:X",
    "Y2:Y",
    "AC2:AC",
    "AD2:AD",
    "AH2:AH",
    "AI2:AI",
    "AM2:AM",
    "AN2:AN",
    "AR2:AR",
    "AS2:AS",
    "AW2:AW",
    "AX2:AX",
    "BB2:BB",
    "BC2:BC",
    "BG2:BG",
    "BH2:BH",
    "BL2:BL",
    "BM2:BM",
    "BQ2:BQ",
    "BR2:BR",
    "BS2:BS",
  ]);
  rangesToFormat.setNumberFormat('"$"#,##0.00');
}

function doPost(e) {
  var id = sheetRegistro.getLastRow();
  var name = e.parameter.name;
  var lastName = e.parameter.apellido;
  var cedula = e.parameter.cedula;
  var address = e.parameter.direccion;
  var city = e.parameter.ciudad;
  var phone = e.parameter.telefono;
  var email = e.parameter.correo;
  var today = new Date();
  var fechaEntrega = e.parameter.fechaEntrega;

  var item1 = e.parameter.productoServicioDatalist1;
  var ref1 = e.parameter.referencia1;
  var cantidad1 = e.parameter.cantidad1;
  var pre1 = e.parameter.precio1;

  var item2 = e.parameter.productoServicioDatalist2;
  var ref2 = e.parameter.referencia2;
  var cantidad2 = e.parameter.cantidad2;
  var pre2 = e.parameter.precio2;

  var item3 = e.parameter.productoServicioDatalist3;
  var ref3 = e.parameter.referencia3;
  var cantidad3 = e.parameter.cantidad3;
  var pre3 = e.parameter.precio3;

  var item4 = e.parameter.productoServicioDatalist4;
  var ref4 = e.parameter.referencia4;
  var cantidad4 = e.parameter.cantidad4;
  var pre4 = e.parameter.precio4;

  var item5 = e.parameter.productoServicioDatalist5;
  var ref5 = e.parameter.referencia5;
  var cantidad5 = e.parameter.cantidad5;
  var pre5 = e.parameter.precio5;

  var item6 = e.parameter.productoServicioDatalist6;
  var ref6 = e.parameter.referencia6;
  var cantidad6 = e.parameter.cantidad6;
  var pre6 = e.parameter.precio6;

  var item7 = e.parameter.productoServicioDatalist7;
  var ref7 = e.parameter.referencia7;
  var cantidad7 = e.parameter.cantidad7;
  var pre7 = e.parameter.precio7;

  var item8 = e.parameter.productoServicioDatalist8;
  var ref8 = e.parameter.referencia8;
  var cantidad8 = e.parameter.cantidad8;
  var pre8 = e.parameter.precio8;

  var item9 = e.parameter.productoServicioDatalist9;
  var ref9 = e.parameter.referencia9;
  var cantidad9 = e.parameter.cantidad9;
  var pre9 = e.parameter.precio9;

  var item10 = e.parameter.productoServicioDatalist10;
  var ref10 = e.parameter.referencia10;
  var cantidad10 = e.parameter.cantidad10;
  var pre10 = e.parameter.precio10;

  var item11 = e.parameter.productoServicioDatalist11;
  var ref11 = e.parameter.referencia11;
  var cantidad11 = e.parameter.cantidad11;
  var pre11 = e.parameter.precio11;

  var item12 = e.parameter.productoServicioDatalist12;
  var ref12 = e.parameter.referencia12;
  var cantidad12 = e.parameter.cantidad12;
  var pre12 = e.parameter.precio12;

  var total1 = pre1 * cantidad1;
  var total2 = pre2 * cantidad2;
  var total3 = pre3 * cantidad3;
  var total4 = pre4 * cantidad4;
  var total5 = pre5 * cantidad5;
  var total6 = pre6 * cantidad6;
  var total7 = pre7 * cantidad7;
  var total8 = pre8 * cantidad8;
  var total9 = pre9 * cantidad9;
  var total10 = pre10 * cantidad10;
  var total11 = pre11 * cantidad11;
  var total12 = pre12 * cantidad12;
  var total =
    total1 +
    total2 +
    total3 +
    total4 +
    total5 +
    total6 +
    total7 +
    total8 +
    total9 +
    total10 +
    total11 +
    total12;

  sheetRegistro.appendRow([
    id,
    today,
    fechaEntrega,
    name,
    lastName,
    cedula,
    address,
    city,
    phone,
    email,
    item1,
    ref1,
    cantidad1,
    pre1,
    total1,
    item2,
    ref2,
    cantidad2,
    pre2,
    total2,
    item3,
    ref3,
    cantidad3,
    pre3,
    total3,
    item4,
    ref4,
    cantidad4,
    pre4,
    total4,
    item5,
    ref5,
    cantidad5,
    pre5,
    total5,
    item6,
    ref6,
    cantidad6,
    pre6,
    total6,
    item7,
    ref7,
    cantidad7,
    pre7,
    total7,
    item8,
    ref8,
    cantidad8,
    pre8,
    total8,
    item9,
    ref9,
    cantidad9,
    pre9,
    total9,
    item10,
    ref10,
    cantidad10,
    pre10,
    total10,
    item11,
    ref11,
    cantidad11,
    pre11,
    total11,
    item12,
    ref12,
    cantidad12,
    pre12,
    total12,
    total,
  ]);
  setCurrencyFormat();
  console.log("se ejecuto doPost");
  return (regTer =
    HtmlService.createTemplateFromFile("registroTerminado").evaluate());
}

var ID = 0,
  TODAY = 1,
  FECHA_ENTREGA = 2,
  NAME = 3,
  LAST_NAME = 4,
  CEDULA = 5,
  ADDRESS = 6,
  CITY = 7,
  PHONE = 8,
  EMAIL = 9,
  ITEM_1 = 10,
  REF_1 = 11,
  CANTIDAD_1 = 12,
  PRE_1 = 13,
  TOTAL_1 = 14;
(ITEM_2 = 15),
  (REF_2 = 16),
  (CANTIDAD_2 = 17),
  (PRE_2 = 18),
  (TOTAL_2 = 19),
  (ITEM_3 = 20),
  (REF_3 = 21),
  (CANTIDAD_3 = 22),
  (PRE_3 = 23),
  (TOTAL_3 = 24),
  (ITEM_4 = 25),
  (REF_4 = 26),
  (CANTIDAD_4 = 27),
  (PRE_4 = 28),
  (TOTAL_4 = 29),
  (ITEM_5 = 30),
  (REF_5 = 31),
  (CANTIDAD_5 = 32),
  (PRE_5 = 33),
  (TOTAL_5 = 34),
  (ITEM_6 = 35),
  (REF_6 = 36),
  (CANTIDAD_6 = 37),
  (PRE_6 = 38),
  (TOTAL_6 = 39),
  (ITEM_7 = 40),
  (REF_7 = 41),
  (CANTIDAD_7 = 42),
  (PRE_7 = 43),
  (TOTAL_7 = 44),
  (ITEM_8 = 45),
  (REF_8 = 46),
  (CANTIDAD_8 = 47),
  (PRE_8 = 48),
  (TOTAL_8 = 49),
  (ITEM_9 = 50),
  (REF_9 = 51),
  (CANTIDAD_9 = 52),
  (PRE_9 = 53),
  (TOTAL_9 = 54),
  (ITEM_10 = 55),
  (REF_10 = 56),
  (CANTIDAD_10 = 57),
  (PRE_10 = 58),
  (TOTAL_10 = 59),
  (ITEM_11 = 60),
  (REF_11 = 61),
  (CANTIDAD_11 = 62),
  (PRE_11 = 63),
  (TOTAL_11 = 64),
  (ITEM_12 = 65),
  (REF_12 = 66),
  (CANTIDAD_12 = 67),
  (PRE_12 = 68),
  (TOTAL_12 = 69),
  (TOTAL = 70);

function crearPdf() {
  var data = sheetRegistro.getDataRange().getValues();
  var lastData = data[data.length - 1];
  console.log(lastData[0]);
  var folderName = "proformas";
  var folder = DriveApp.getFoldersByName(folderName);
  if (folder.hasNext()) {
    folder = folder.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  id = lastData[ID];
  today = lastData[TODAY].toLocaleDateString("en-GB");
  fechaEntrega = lastData[FECHA_ENTREGA].toLocaleDateString("en-GB");
  name = lastData[NAME].toUpperCase();
  lastName = lastData[LAST_NAME].toUpperCase();
  cedula = lastData[CEDULA].toLocaleString("de-DE");
  address = lastData[ADDRESS].toUpperCase();
  city = lastData[CITY].toUpperCase();
  phone = lastData[PHONE];
  email = lastData[EMAIL].toUpperCase();
  item1 = lastData[ITEM_1].toUpperCase();
  ref1 = lastData[REF_1];
  cantidad1 = lastData[CANTIDAD_1];
  pre1 = lastData[PRE_1].toLocaleString();
  total1 = lastData[TOTAL_1].toLocaleString();
  item2 = lastData[ITEM_2].toUpperCase();
  ref2 = lastData[REF_2];
  cantidad2 = lastData[CANTIDAD_2];
  pre2 = lastData[PRE_2].toLocaleString();
  total2 = lastData[TOTAL_2].toLocaleString();
  item3 = lastData[ITEM_3].toUpperCase();
  ref3 = lastData[REF_3];
  cantidad3 = lastData[CANTIDAD_3];
  pre3 = lastData[PRE_3].toLocaleString();
  total3 = lastData[TOTAL_3].toLocaleString();
  item4 = lastData[ITEM_4].toUpperCase();
  ref4 = lastData[REF_4];
  cantidad4 = lastData[CANTIDAD_4];
  pre4 = lastData[PRE_4].toLocaleString();
  total4 = lastData[TOTAL_4].toLocaleString();
  item5 = lastData[ITEM_5].toUpperCase();
  ref5 = lastData[REF_5];
  cantidad5 = lastData[CANTIDAD_5];
  pre5 = lastData[PRE_5].toLocaleString();
  total5 = lastData[TOTAL_5].toLocaleString();
  item6 = lastData[ITEM_6].toUpperCase();
  ref6 = lastData[REF_6];
  cantidad6 = lastData[CANTIDAD_6];
  pre6 = lastData[PRE_6].toLocaleString();
  total6 = lastData[TOTAL_6].toLocaleString();
  item7 = lastData[ITEM_7].toUpperCase();
  ref7 = lastData[REF_7];
  cantidad7 = lastData[CANTIDAD_7];
  pre7 = lastData[PRE_7].toLocaleString();
  total7 = lastData[TOTAL_7].toLocaleString();
  item8 = lastData[ITEM_8].toUpperCase();
  ref8 = lastData[REF_8];
  cantidad8 = lastData[CANTIDAD_8];
  pre8 = lastData[PRE_8].toLocaleString();
  total8 = lastData[TOTAL_8].toLocaleString();
  item9 = lastData[ITEM_9].toUpperCase();
  ref9 = lastData[REF_9];
  cantidad9 = lastData[CANTIDAD_9];
  pre9 = lastData[PRE_9].toLocaleString();
  total9 = lastData[TOTAL_9].toLocaleString();
  item10 = lastData[ITEM_10].toUpperCase();
  ref10 = lastData[REF_10];
  cantidad10 = lastData[CANTIDAD_10];
  pre10 = lastData[PRE_10].toLocaleString();
  total10 = lastData[TOTAL_10].toLocaleString();
  item11 = lastData[ITEM_11].toUpperCase();
  ref11 = lastData[REF_11];
  cantidad11 = lastData[CANTIDAD_11];
  pre11 = lastData[PRE_11].toLocaleString();
  total11 = lastData[TOTAL_11].toLocaleString();
  item12 = lastData[ITEM_12].toUpperCase();
  ref12 = lastData[REF_12];
  cantidad12 = lastData[CANTIDAD_12];
  pre12 = lastData[PRE_12].toLocaleString();
  total12 = lastData[TOTAL_12].toLocaleString();
  total = lastData[TOTAL].toLocaleString();

  const variable = [
    item12,
    item11,
    item10,
    item9,
    item8,
    item7,
    item6,
    item5,
    item4,
    item3,
    item2,
    item1,
  ];

  let contador = 12 - variable.findIndex((el) => el != "");

  var archivoPlantilla = DriveApp.getFileById(ID_PLANTILLA[contador]);
  var copyPlantilla = archivoPlantilla.makeCopy();
  folder.addFile(copyPlantilla);
  copyPlantilla.setName(`proforma_${id}_${name}_${lastName}`);

  var idArchivoCopia = copyPlantilla.getId();

  var doc = DocumentApp.openById(idArchivoCopia);
  var txt = doc.getBody();
  txt.replaceText("{{ID}}", id);
  txt.replaceText("{{FECHA}}", today);
  txt.replaceText("{{ENTREGA}}", fechaEntrega);
  txt.replaceText("{{NOMBRE}}", name);
  txt.replaceText("{{APELLIDO}}", lastName);
  txt.replaceText("{{CEDULA}}", cedula);
  txt.replaceText("{{DIRECCION}}", address);
  txt.replaceText("{{CIUDAD}}", city);
  txt.replaceText("{{TELEFONO}}", phone);
  txt.replaceText("{{EMAIL}}", email);
  txt.replaceText("{{item1}}", item1);
  txt.replaceText("{{ref1}}", ref1);
  txt.replaceText("{{pre1}}", pre1);
  txt.replaceText("{{cantidad1}}", cantidad1);
  txt.replaceText("{{item2}}", item2);
  txt.replaceText("{{ref2}}", ref2);
  txt.replaceText("{{pre2}}", pre2);
  txt.replaceText("{{cantidad2}}", cantidad2);
  txt.replaceText("{{item3}}", item3);
  txt.replaceText("{{ref3}}", ref3);
  txt.replaceText("{{pre3}}", pre3);
  txt.replaceText("{{cantidad3}}", cantidad3);
  txt.replaceText("{{item4}}", item4);
  txt.replaceText("{{ref4}}", ref4);
  txt.replaceText("{{pre4}}", pre4);
  txt.replaceText("{{cantidad4}}", cantidad4);
  txt.replaceText("{{item5}}", item5);
  txt.replaceText("{{ref5}}", ref5);
  txt.replaceText("{{pre5}}", pre5);
  txt.replaceText("{{cantidad5}}", cantidad5);
  txt.replaceText("{{item6}}", item6);
  txt.replaceText("{{ref6}}", ref6);
  txt.replaceText("{{pre6}}", pre6);
  txt.replaceText("{{cantidad6}}", cantidad6);
  txt.replaceText("{{item7}}", item7);
  txt.replaceText("{{ref7}}", ref7);
  txt.replaceText("{{pre7}}", pre7);
  txt.replaceText("{{cantidad7}}", cantidad7);
  txt.replaceText("{{item8}}", item8);
  txt.replaceText("{{ref8}}", ref8);
  txt.replaceText("{{pre8}}", pre8);
  txt.replaceText("{{cantidad8}}", cantidad8);
  txt.replaceText("{{item9}}", item9);
  txt.replaceText("{{ref9}}", ref9);
  txt.replaceText("{{pre9}}", pre9);
  txt.replaceText("{{cantidad9}}", cantidad9);
  txt.replaceText("{{item10}}", item10);
  txt.replaceText("{{ref10}}", ref10);
  txt.replaceText("{{pre10}}", pre10);
  txt.replaceText("{{cantidad10}}", cantidad10);
  txt.replaceText("{{item11}}", item11);
  txt.replaceText("{{ref11}}", ref11);
  txt.replaceText("{{pre11}}", pre11);
  txt.replaceText("{{cantidad11}}", cantidad11);
  txt.replaceText("{{item12}}", item12);
  txt.replaceText("{{ref12}}", ref12);
  txt.replaceText("{{pre12}}", pre12);
  txt.replaceText("{{cantidad12}}", cantidad12);
  txt.replaceText("{{total1}}", total1);
  txt.replaceText("{{total2}}", total2);
  txt.replaceText("{{total3}}", total3);
  txt.replaceText("{{total4}}", total4);
  txt.replaceText("{{total5}}", total5);
  txt.replaceText("{{total6}}", total6);
  txt.replaceText("{{total7}}", total7);
  txt.replaceText("{{total8}}", total8);
  txt.replaceText("{{total9}}", total9);
  txt.replaceText("{{total10}}", total10);
  txt.replaceText("{{total11}}", total11);
  txt.replaceText("{{total12}}", total12);
  txt.replaceText("{{TOTAL}}", total);

  /* txt.replaceText("{{total}}", total) */

  doc.saveAndClose();

  /* var blob = archivoPlantilla.getBlob().setName(`proforma_${id}.pdf`);
    folder.createFile(blob) */

  /*   var html = HtmlService.createTemplateFromFile("plantillaPdf").evaluate();
    var pdf = html.getAs("application/pdf").setName(`proforma_${id}.pdf`);
    folder.createFile(pdf); */
}

function cargarPendientes() {
  return HtmlService.createTemplateFromFile("trabajosPendientes").evaluate();
}
