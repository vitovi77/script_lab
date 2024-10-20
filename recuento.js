$("#filtro-Lote").on("click", () => tryCatch(FiltroLote));
$("#filtro-Zona").on("click", () => tryCatch(FiltroZona));
$("#filtro-Color").on("click", () => tryCatch(FiltroColor));
$("#filtro-Calidad").on("click", () => tryCatch(FiltroCalidad));
$("#filtro-Articulo").on("click", () => tryCatch(FiltroArticulo));
$("#Desactivar-Filtros").on("click", () => tryCatch(DesactivarFiltros));
$("#añadir").on("click", () => tryCatch(Enviardatos));
$("#borrar-datos").on("click", () => tryCatch(Borrardatos));

async function Enviardatos() {
  await Excel.run(async (context) => {
    var mihoja = context.workbook.worksheets.getItem("Buscador");

    var mirango = mihoja.getUsedRange();
    mirango.load("rowCount");
    await context.sync();
    var nfila = mirango.rowCount;

    //última celda llena de la columna A
    var celda1 = mihoja.getRange("A" + nfila);
    celda1.load("values");
    await context.sync();
    //Valor de la última celda llena de la columna A
    var valor = celda1.values[0][0];
    //Agregamos una fila
    var nfila = nfila + 1;
    //Sumamos 1 al valor de la columna N° (número conswcutivo)
    var valor = valor + 1;
    //Le damos nuevo valor - consecutivo - a la celda siguiente de la columna N°
    mihoja.getRange("A" + nfila).values = valor;
    //Llenamos las otras celdas de la fila con los datos del formulario
    mihoja.getRange("A" + nfila).values = document.getElementById("Lote").value;
    mihoja.getRange("C" + nfila).values = document.getElementById("Cantidad").value;
    mihoja.getRange("B" + nfila).values = document.getElementById("Zona").value;
    mihoja.getRange("F" + nfila).values = document.getElementById("Color").value;
    mihoja.getRange("E" + nfila).values = document.getElementById("Articulo").value;

    mihoja.getRange("D" + nfila).values = document.getElementById("Producto").value;
    mihoja.getRange("D" + nfila).values = document.getElementById("Producto").value;

    mihoja.getRange("G" + nfila).values = document.getElementById("Calidad").value;
    mihoja.getRange("J" + nfila).values = document.getElementById("Notas").value;

    await context.sync();
  });
}

async function Borrardatos() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    await context.sync();
    document.getElementById("Lote").value = "*";
    document.getElementById("Zona").value = "";
    document.getElementById("Articulo").value = "";
    document.getElementById("Color").value = "";
    document.getElementById("Calidad").value = "";
    document.getElementById("Cantidad").value = "";
    document.getElementById("Producto").value = "";
    document.getElementById("Notas").value = "";
    await context.sync();
    await context.sync();
  });
}

async function FiltroLote() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const farmData = sheet.getUsedRange();

    //valor del lote es el de la cajabox
    let lote = document.getElementById("Lote").value;

    sheet.autoFilter.apply(farmData, 0, {
      criterion1: lote,
      filterOn: Excel.FilterOn.custom
    });

    await context.sync();
  });
}

async function FiltroArticulo() {
  await Excel.run(async (context) => {
    var mihoja = context.workbook.worksheets.getItem("Buscador");
    var mirango = mihoja.getUsedRange();
    mirango.load("rowCount");
    await context.sync();
    var nfila = mirango.rowCount;

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const farmData = sheet.getUsedRange();
    let articulo = document.getElementById("Articulo").value;
    if (articulo === "006950") document.getElementById("Producto").value = "Endurance 100x100 7mm st6 / 006950";
    else if (articulo === "007062") document.getElementById("Producto").value = "Extreme 100x100 7mm ST6 / 007062";
    else if (articulo === "006116") document.getElementById("Producto").value = "PAVIGYM Extreme 100x100 7mm / 006116";
    else if (articulo === "006946")
      document.getElementById("Producto").value = "Endurance sin logo 100x100 7mm st6 / 006946";
    else if (articulo === "006956") document.getElementById("Producto").value = "Performance+ 100x100 5mm st6 / 006956";
    else if (articulo === "005991")
      document.getElementById("Producto").value = "Cycle free weight SS 3 mm v20 st5 / 005991";
    else if (articulo === "006995")
      document.getElementById("Producto").value = "Endurance Cobertura 100x100 3mm ST5 / 006995";
    else if (articulo === "007141") document.getElementById("Producto").value = "Motion 100x100 9mm ST6 / 007141";
    else if (articulo === "006945")
      document.getElementById("Producto").value = "PAVIGYM  Endurance 100x100 7mm / 006945";
    else if (articulo === "007159")
      document.getElementById("Producto").value = "PAVIGYM  Endurance S&S 100x100 22mm / 007159";
    else if (articulo === "007846")
      document.getElementById("Producto").value = "ACOUSTIC BigJag XL Functional 100x100 55mm / 007846";
    else if (articulo === "007845")
      document.getElementById("Producto").value = "ACOUSTIC BigJag L Functional 100x100 42,5mm / 00784";
    else if (articulo === "006299")
      document.getElementById("Producto").value = "PAVIGYM Extreme S&S 100x100cm 22mm / 006299";
    else if (articulo === "000087") document.getElementById("Producto").value = "PAVIGYM Extreme 90x90 7mm / 000087";
    else if (articulo === "005494")
      document.getElementById("Producto").value = "PAVIGYM  Endurance 90x90 7mm v20 / 005494";
    else if (articulo === "006475")
      document.getElementById("Producto").value = "PAVIGYM Performance+ 90x90cm 5mm v20 / 006475";
    else if (articulo === "000019") document.getElementById("Producto").value = "PAVIGYM  Endurance 90x90 7mm / 000019";
    else if (articulo === "000788")
      document.getElementById("Producto").value = "PAVIGYM  Endurance S&S 90x90 22mm / 000788";
    else if (articulo === "004573") document.getElementById("Producto").value = "Weightlifting PVC 22mm st6 / 004573";
    else if (articulo === "006482")
      document.getElementById("Producto").value = "Performance+ 90x90 5mm v20 st6 / 006482";
    else if (articulo === "002227") document.getElementById("Producto").value = "Cycle free weight 7mm st6 / 002227";
    else if (articulo === "006467")
      document.getElementById("Producto").value = "PAVIGYM Functional 90x90cm 20mm v20 / 006467";
    else if (articulo === "005488")
      document.getElementById("Producto").value = "Cycle free weight 7mm v20 st6 / 005488";
    else if (articulo === "006937")
      document.getElementById("Producto").value = "PAVIGYM Performance+ 100x100cm 5mm / 006937";
    else if (articulo === "000019") document.getElementById("Producto").value = "PAVIGYM  Endurance 90x90 7mm / 000019";
    else if (articulo === "007444")
      document.getElementById("Producto").value = "PAVIGYM Solid Pro 100x100 30mm / 007444";
    else if (articulo === "007047")
      document.getElementById("Producto").value = "ACOUSTIC BigJag Endurance 100x100 90mm / 007047";
    else if (articulo === "006745")
      document.getElementById("Producto").value = "Functional Cobertura 90x90 5mm v20 st6 / 006745";
    else if (articulo === "007150")
      document.getElementById("Producto").value = "PAVIGYM Motion 100x100 9mm without logo / 007150";
    else if (articulo === "007147") document.getElementById("Producto").value = "PAVIGYM Motion 100x100 9mm / 007147";
    else if (articulo === "006298") document.getElementById("Producto").value = "Extreme S&S 100x100 22mm st6 / 006298";
    else if (articulo === "000088")
      document.getElementById("Producto").value = "PAVIGYM Extreme S&S 90x90 22mm / 000088";
    else if (articulo === "005968")
      document.getElementById("Producto").value = "PAVIGYM  Endurance S&S 90x90 22mm v20 / 005968";
    else if (articulo === "007045")
      document.getElementById("Producto").value = "ACOUSTIC BigJag Endurance 100x100 50mm / 007045";
    else if (articulo === "007048")
      document.getElementById("Producto").value = "PAVIBASIC SOLID Endurance 100x100 40mm ST4 / 007048";
    else if (articulo === "005634")
      document.getElementById("Producto").value = "PAVIGYM Performance+ 90x90cm 5mm / 005634";
    else if (articulo === "006443") document.getElementById("Producto").value = "PAVIPLAY (50x50cm 18mm) / 006443";
    else if (articulo === "007694")
      document.getElementById("Producto").value = "PAVIBASIC Cross Training 100x100 30mm reduced interlocking /";
    else if (articulo === "006204")
      document.getElementById("Producto").value = "PAVIBASIC Cross Training 96x96 30mm / 006204-BLACK";
    else if (articulo === "007024")
      document.getElementById("Producto").value = "Extreme Cobertura 100x100 6mm ST4 / 007024";
    else if (articulo === "007475")
      document.getElementById("Producto").value = "Cobertura Solid Pro 100x100 6mm ST4 / 007475";
    else if (articulo === "007488")
      document.getElementById("Producto").value = "Cobertura Extreme 100x100 3mm PYC + BIGJAG ST5 / 007488";
    else if (articulo === "006289") document.getElementById("Producto").value = "Vertical 105x205 7mm st6 / 006289";
    else if (articulo === "007049")
      document.getElementById("Producto").value = "PAVIBASIC SOLID Extreme 100x100 40mm ST4 / 007049";
    else if (articulo === "002398")
      document.getElementById("Producto").value = "Pavigym reciclado 7mm independiente st6 / 002398";
    else if (articulo === "006292")
      document.getElementById("Producto").value = "Extreme 105x105 7mm CASTER st6 / 006292";
    else if (articulo === "006148")
      document.getElementById("Producto").value = "Cycle free weight 7mm v20 Molde E st6 / 006148";
    else if (articulo === "006135") document.getElementById("Producto").value = "Cycle free weight 7mm PI ST6 / 006135";
    else if (articulo === "006484")
      document.getElementById("Producto").value = "Functional 90x90 20mm v20 st6 / 006484";
    else if (articulo === "007162")
      document.getElementById("Producto").value = "PAVIGYM  Endurance S&S 100x100 22mm without logo / 007162";
    else if (articulo === "002400")
      document.getElementById("Producto").value = "Pavigym extreme S&S (90x90) 22mm st6 / 002400";
    else if (articulo === "002387")
      document.getElementById("Producto").value = "Cycle free weight  SS 22 mm st6 / 002387";
    else if (articulo === "007141") document.getElementById("Producto").value = "Motion 100x100 9mm ST6 / 007141";
    else if (articulo === "007163")
      document.getElementById("Producto").value = "Endurance S&S 100x100 22mm ST6 / 007163";
    else if (articulo === "001215")
      document.getElementById("Producto").value = "Pavigym reciclado 7mm independiente st4 / 001215";
    else if (articulo === "007139") document.getElementById("Producto").value = "Motion sin logo 100x100 st6 / 007139";
    else if (articulo === "006155")
      document.getElementById("Producto").value = "PAVIBASIC Cross Training 96x96 20mm / 006155";
    else if (articulo === "006176")
      document.getElementById("Producto").value = "PAVIBASIC Cross Training 96x96cm 25mm / 006176";
    else if (articulo === "006236") document.getElementById("Producto").value = "PAVIBASIC Cardio 96x96 10mm / 006236";
    else if (articulo === "006945")
      document.getElementById("Producto").value = "PAVIGYM  Endurance 100x100 7mm / 006945";
    else if (articulo === "007013")
      document.getElementById("Producto").value = "ACOUSTIC BigJag Extreme 100x100 70mm / 007013";
    else if (articulo === "007017")
      document.getElementById("Producto").value = "ACOUSTIC BigJag CT 100x100 90mm / 007017";
    else if (articulo === "006241")
      document.getElementById("Producto").value = "Cycle free weight  3,5 mm ​​PI st5 / 006241";
    else if (articulo === "006047") document.getElementById("Producto").value = "Motion Hard XL 9mm st6 / 006047";
    else if (articulo === "002415") document.getElementById("Producto").value = "Motion Extreme XL 9mm st6 / 002415";
    else if (articulo === "000001") document.getElementById("Producto").value = "PAVIGYM  Motion 90x90 9mm / 000001";
    else if (articulo === "007847")
      document.getElementById("Producto").value = "ACOUSTIC BigJag XXL Functional 100x100 80mm / 007847";
    else if (articulo === "007046")
      document.getElementById("Producto").value = "ACOUSTIC BigJag Endurance 100x100 70mm / 007046";
    else if (articulo === "007023")
      document.getElementById("Producto").value = "Extreme Cobertura 100x100 3mm ST5 / 007023";
    else if (articulo === "006487")
      document.getElementById("Producto").value = "Weightlifting PVC 22mm v20 st6 / 006487";
    else if (articulo === "007746")
      document.getElementById("Producto").value = "Motion 100x100 9mm ST6 (Solid Pro Surface) / 007746";
    else if (articulo === "003163")
      document.getElementById("Producto").value = "ENDURANCE NO CONFORME 6MM st6 / 003163--NQ";
    else if (articulo === "006961")
      document.getElementById("Producto").value = "Performance+ sin logo 100x100 5mm st6 / 006961";
    else if (articulo === "007395")
      document.getElementById("Producto").value = "Endurance no conforme 6mm st6 (O100) / 007395";
    else if (articulo === "007865")
      document.getElementById("Producto").value = "PAVIGYM Endurance Bfl-s1 100x100 7mm / 007865";
    else if (articulo === "006468")
      document.getElementById("Producto").value = "PAVIGYM  Motion 90x90 9mm v20 / 006468";
    else if (articulo === "006949")
      document.getElementById("Producto").value = "PAVIGYM  Endurance 100x100 7mm without logo / 006949";
    else if (articulo === "006121")
      document.getElementById("Producto").value = "PAVIGYM Endurance 105x105 7mm ST6 / 006121";
    else if (articulo === "004547")
      document.getElementById("Producto").value = "PAVIGYM Weightlifting 90x90cm 22mm / 004547";
    else if (articulo === "006957") document.getElementById("Producto").value = "Performance+ 100x100 5mm st4 / 006957";
    else if (articulo === "007012")
      document.getElementById("Producto").value = "ACOUSTIC BigJag Extreme 100x100 50mm / 007012";
    else if (articulo === "005516") document.getElementById("Producto").value = "Functional 90x90cm 20mm st6 / 005516";
    else if (articulo === "007473")
      document.getElementById("Producto").value = "PAVIGYM Solid Pro 100x100 20mm / 007473";
    else if (articulo === "007848")
      document.getElementById("Producto").value = "ACOUSTIC BigJag L Extreme 100x100 42,5mm / 007848";
    else if (articulo === "007849")
      document.getElementById("Producto").value = "ACOUSTIC BigJag XL Extreme 100x100 55mm / 007849";
    else if (articulo === "007850")
      document.getElementById("Producto").value = "ACOUSTIC BigJag XXL Extreme 100x100 80mm / 007850";
    else if (articulo === "007110")
      document.getElementById("Producto").value = "PAVIGYM Weightlifting 100x100cm 22mm / 007110";
    else if (articulo === "007111")
      document.getElementById("Producto").value = "Weightlifting PVC 100x100 22mm ST6 / 007111";
    else if (articulo === "007014")
      document.getElementById("Producto").value = "ACOUSTIC BigJag Extreme 100x100 90mm / 007014";
    else if (articulo === "007022")
      document.getElementById("Producto").value = "Endurance Cobertura sin logo 100x100 3mm ST5 / 007022";
    else if (articulo === "007476")
      document.getElementById("Producto").value = "Cobertura Solid Pro 100x100 2,5mm ST5 / 007476";
    else if (articulo === "007778")
      document.getElementById("Producto").value = "Cobertura Solid Pro 100x100 2,25mm PARA MOTION Y BIGJAG ST5 ";
    else if (articulo === "007018")
      document.getElementById("Producto").value = "PAVIBASIC CT 100x100 40mm ST4 sin cobertura / 007018";
    else if (articulo === "007016")
      document.getElementById("Producto").value = "ACOUSTIC BigJag CT 100x100 70mm / 007016";
    else if (articulo === "007795") document.getElementById("Producto").value = "Motion 100x100 9mm st4 / 007795";
    else if (articulo === "007816")
      document.getElementById("Producto").value = "PAVIBASIC CT 100x100 30mm antishock / 007816";
    else if (articulo === "002396")
      document.getElementById("Producto").value = "Pavigym reciclado cobertura SS 3,5mm st5 / 002396";
    else mihoja.getRange("D" + nfila).values = document.getElementById("Producto").value;

    sheet.autoFilter.apply(farmData, 4, {
      criterion1: articulo,
      filterOn: Excel.FilterOn.custom
    });

    await context.sync();
  });
}

async function FiltroColor() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const farmData = sheet.getUsedRange();
    let color = document.getElementById("Color").value;
    sheet.autoFilter.apply(farmData, 5, {
      criterion1: color,
      filterOn: Excel.FilterOn.custom
    });

    await context.sync();
  });
}

async function FiltroCalidad() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const farmData = sheet.getUsedRange();
    let color = document.getElementById("Calidad").value;
    sheet.autoFilter.apply(farmData, 6, {
      criterion1: color,
      filterOn: Excel.FilterOn.custom
    });

    await context.sync();
  });
}

async function FiltroZona() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const farmData = sheet.getUsedRange();
    let zona = document.getElementById("Zona").value;
    sheet.autoFilter.apply(farmData, 1, {
      criterion1: zona,
      filterOn: Excel.FilterOn.custom
    });

    await context.sync();
  });
}

async function DesactivarFiltros() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.remove();
    await context.sync();
    document.getElementById("Lote").value = "*";
    document.getElementById("Zona").value = "";
    document.getElementById("Articulo").value = "";
    document.getElementById("Color").value = "";
    document.getElementById("Calidad").value = "";
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
