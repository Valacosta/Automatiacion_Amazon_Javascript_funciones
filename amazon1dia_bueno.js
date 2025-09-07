
'declaracion de librerias :)'
import puppeteer from "puppeteer";
import { createInterface } from "node:readline";
import fs from "node:fs";
import Excel from "exceljs";

//Se define la función de run, ya que sera la que corre la funcion principal'
let hoy_aux;
async function run() {

    const rl = createInterface({
    input: process.stdin,
    output: process.stdout,
  });

 let cancelaciones=[];

 //Esta funcion nos ayudara a pausar al iniciar sesion y dar enter para continuar '
  function obtenerNombre() {
    return new Promise((resolve) => {
      rl.question("Inicie sesión y de enter por favor: ", (nombre) => {
        
        rl.close();
        resolve(nombre);
      });
    });
  }

  //Esta funcion nos ayudara a pausar y dar un tiempo de espera un delay'
  function delay(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  //Esta funcion nos ayuda a inicializar el navegador '

  async function editarExcel(rutaArchivo, numerocel, datitos,hojaactivar) {
      
      try {
    
          console.log(`Intentando leer el archivo para editar: ${rutaArchivo}`); // Agregar log
          const workbook = new Excel.Workbook();
          await workbook.xlsx.readFile(rutaArchivo);
          const worksheet = workbook.getWorksheet(hojaactivar);
          worksheet.getCell('A1').value = 'Nuevo Valor';
          worksheet.addRow(datitos);
  
          await workbook.xlsx.writeFile(rutaArchivo);
          console.log('Archivo Excel editado exitosamente.');
  
      } catch (error) {
          console.error('Error al editar el archivo Excel:', error);
      }
  }
  function compararCadenas(a, b) {
      if (a === b) {
        console.log("✅ Las cadenas son exactamente iguales.");
      } else {
        console.log("❌ Las cadenas NO son iguales.");
        console.log("a:", a, "| length:", a.length);
        console.log("b:", b, "| length:", b.length);
  
        const minLen = Math.min(a.length, b.length);
        for (let i = 0; i < minLen; i++) {
          if (a[i] !== b[i]) {
            console.log(
              `Diferencia en posición ${i}: '${a[i]}' (charCode ${a.charCodeAt(
                i
              )}) vs '${b[i]}' (charCode ${b.charCodeAt(i)})`
            );
          }
        }
  
        if (a.length !== b.length) {
          console.log("⚠️ Las longitudes son diferentes.");
          for (let i = minLen; i < a.length; i++) {
            console.log(`Extra en a: '${a[i]}' (charCode ${a.charCodeAt(i)})`);
          }
          for (let i = minLen; i < b.length; i++) {
            console.log(`Extra en b: '${b[i]}' (charCode ${b.charCodeAt(i)})`);
          }
        }
      }
    }
    function crearArchivoExcel(filepath) {
      const workbook = new Excel.Workbook();
      const worksheet = workbook.addWorksheet("1_Guia");
      const worksheet2 = workbook.addWorksheet("2_Guias");
  
      // Agregar los encabezados
      worksheet.addRow([
        "NO. VENTA",
        "LLANTAS A DESPACHAR",
        "DESCRIPCION",
        "NUMERO DE PAQUETES",
        "PRECIO POR UNIDAD",
        "ENVIO",
        "VENTA TOTAL",
        "PU S/IVA",
        "NOTA DE VENTA",
        "GUIA",
        "CLIENTE",
        "NUMERO DE TELEFONO",
        "CLAVE",
        "LLANTAS  POR PUBLICACION",
        "LLANTAS POR VENTA",
      ]);
      worksheet2.addRow([
        "NO. VENTA",
        "LLANTAS A DESPACHAR",
        "DESCRIPCION",
        "NUMERO DE PAQUETES",
        "PRECIO POR UNIDAD",
        "ENVIO",
        "VENTA TOTAL",
        "PU S/IVA",
        "NOTA DE VENTA",
        "GUIA",
        "CLIENTE",
        "NUMERO DE TELEFONO",
        "CLAVE",
        "LLANTAS  POR PUBLICACION",
        "LLANTAS POR VENTA",
      ]);
  
      // Establecer el ancho de las columnas
      const columnasAncho = {
        A: 23,
        B: 10,
        C: 45,
        D: 10,
        E: 17,
        F: 9,
        G: 13,
        H: 14,
        I: 12,
        J: 12,
        K: 35,
        L: 15,
        M: 16,
      };
      for (const columna in columnasAncho) {
        worksheet.getColumn(columna).width = columnasAncho[columna];
        worksheet2.getColumn(columna).width = columnasAncho[columna];
      }
  
      // Estilizar la primera fila (encabezados)
      const startColumn = "A";
      const startRow = 1;
      const endColumn = "O";
      const endRow = 1;
      const headerBackgroundColor = "1188FF";
  
      for (
        let colCode = startColumn.charCodeAt(0);
        colCode <= endColumn.charCodeAt(0);
        colCode++
      ) {
        let columnName = String.fromCharCode(colCode);
        let cell = worksheet.getCell(`${columnName}${startRow}`);
        let cell2 = worksheet2.getCell(`${columnName}${startRow}`);
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: headerBackgroundColor },
        };
        cell2.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: headerBackgroundColor },
        };
        cell.alignment = {
          wrapText: true,
          vertical: "middle",
          horizontal: "center",
        };
        cell2.alignment = {
          wrapText: true,
          vertical: "middle",
          horizontal: "center",
        };
      }
  
      // Estilizar celdas específicas
      const celdasEstilo = {
        B1: { backgroundColor: "FFFF00" },
        G1: { backgroundColor: "FFA63C" },
        L1: { backgroundColor: "B6B3B2" },
      };
      for (const cellAddress in celdasEstilo) {
        var cell = worksheet.getCell(cellAddress);
        var cell2 = worksheet2.getCell(cellAddress);
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: celdasEstilo[cellAddress].backgroundColor },
        };
        cell2.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: celdasEstilo[cellAddress].backgroundColor },
        };
      }
  
      workbook.xlsx.writeFile(filepath);
      console.log(`Archivo Excel '${filepath}' creado.`);
    }
  function regresames(mes)
  {
      if (mes == 1) {
    mes = "ene";
    } else if (mes == 2) {
        mes = "feb";
    } else if (mes == 3) {
        mes = "mar";
    } else if (mes == 4) {
        mes = "abr";
    } else if (mes == 5) {
        mes = "may";
    } else if (mes == 6) {
        mes = "jun";
    } else if (mes == 7) {
        mes = "jul";
    } else if (mes == 8) {
        mes = "ago";
    } else if (mes == 9) {
        mes = "sep";
    } else if (mes == 10) {
        mes = "oct";
    } else if (mes == 11) {
        mes = "nov";
    } else if (mes == 12) {
        mes = "dic";
    }
    return mes;
  }

  async function obtenerTextoPorXPath(page, xpath_selector) {
  const texto = await page.evaluate((xpath) => {
    const el = document.evaluate(
      xpath,
      document,
      null,
      XPathResult.FIRST_ORDERED_NODE_TYPE,
      null
    ).singleNodeValue;
    return el ? el.textContent.trim() : null;
  }, xpath_selector);

  return texto;
}

function extraerNumerosDeString(str) {
  // 1. Verificar si la entrada es una cadena de texto válida
  if (typeof str !== 'string') {
    console.error("Error: La entrada debe ser una cadena de texto.");
    return null;
  }

  // 2. Usar una expresión regular para encontrar todos los dígitos
  // La expresión regular `/\d/g` busca cualquier dígito (0-9) globalmente.
  const numerosEncontrados = str.match(/\d/g);

  // 3. Verificar si se encontraron dígitos
  if (!numerosEncontrados || numerosEncontrados.length === 0) {
    console.log("No se encontraron números en la cadena.");
    return null;
  }

  // 4. Unir los dígitos en una sola cadena
  const cadenaDeNumeros = numerosEncontrados.join('');

  // 5. Convertir la cadena de números a un número entero
  const numeroEntero = parseInt(cadenaDeNumeros, 10);

  // 6. Retornar el número
  return numeroEntero;
}

//se declara la funcion de clickear un elemento dada la agina y el xpath
async function clickElementoPorXPath(page, xpath_selector) {
  await page.evaluate((xpath) => {
    var el = document.evaluate(
      xpath,
      document,
      null,
      XPathResult.FIRST_ORDERED_NODE_TYPE,
      null
    ).singleNodeValue;
    if (el) el.click();
  }, xpath_selector);
} 
//Se inicializa el driver
    //Aqui se pone laruta del perfilque se quiere abrir'
    const userDataDir ="C:\\Users\\LENOVO GIL\\AppData\\Local\\Google\\Chrome\\User Data\\Default";
    //Aqui se pone el la ruta al ejecutale del chrome'
    const chromeExecutablePath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";
      const browser = await puppeteer.launch({
        headless: false, // Ejecutar con interfaz gráfica para ser menos detectable
        userDataDir: userDataDir,
        executablePath: chromeExecutablePath,
        args: [
          "--start-maximized",
          "--disable-blink-features=AutomationControlled",
        ],
        ignoreDefaultArgs: ["--enable-automation"],
      });
      const page = await browser.newPage();
    
      // Simular propiedades del navegador para ser menos detectable
      await page.setUserAgent(
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
      );
      await page.evaluateOnNewDocument(() => {
        Object.defineProperty(navigator, "webdriver", {
          get: () => undefined,
        });
      });
    
      await page.goto(
        "https://sellercentral.amazon.com.mx/orders-v3/mfn/shipped/selfship?ref_=xx_swlang_head_xx&mons_sel_locale=es_MX&languageSwitched=1&page=1"
      ); // Navega a tu sitio objetivo
      await obtenerNombre();

  //Se obtienen las 2 fechas a utilizar del dia de hoy 
  const hoy = new Date();
  const año = hoy.getFullYear(); // Obtiene el año (ej: 2025)
  let mes_temp = hoy.getMonth() + 1; // Obtiene el mes (0-11, por eso le sumamos 1)
  const dia = hoy.getDate();
  const hour = hoy.getHours();
  const minute = hoy.getMinutes();
  const milisec = hoy.getMilliseconds();
  const hoy_completo2 = dia + "/" + mes_temp + "/" + año;
  let mes = regresames(mes_temp);
  var hoy_completo = dia + " de " + mes + " de " + año;
  hoy_aux = hoy_completo;
  //console.log(`Fecha actual es: ${hoy_completo}\nFecha actual2: ${hoy_completo2}`);

  //Se obtienen las 2 fechas a utilizr deldia de ayer 
  const ayer = new Date(hoy);
  ayer.setDate(hoy.getDate() - 1); // Restamos 1 día
  const año_ayer = ayer.getFullYear(); // Obtiene el año (ej: 2025)
  let mes_ayer_temp = ayer.getMonth() + 1; // Obtiene el mes (0-11, por eso le sumamos 1)
  const dia_ayer = ayer.getDate();
  const ayer_completo2 = dia_ayer + "/" + mes_ayer_temp + "/" + año_ayer;
  let mes_ayer = regresames(mes_ayer_temp);
  var ayer_completo = dia_ayer + " de " + mes_ayer + " de " + año_ayer;
  //console.log(`Fecha ayer es: ${ayer_completo}\nFecha ayer2: ${ayer_completo2}`);

  // Se obtienen las 2 fechas del da de antier 
  const antier = new Date(hoy);
  antier.setDate(hoy.getDate() - 2); // Restamos 1 día
  const año_antier = antier.getFullYear(); // Obtiene el año (ej: 2025)
  let mes_antier_temp = antier.getMonth() + 1; // Obtiene el mes (0-11, por eso le sumamos 1)
  const dia_antier = antier.getDate();
  const antier_completo2 = dia_antier + "/" + mes_antier_temp + "/" + año_antier;
  let mes_antier = regresames(mes_antier_temp);
  const antier_completo = dia_antier + " de " + mes_antier + " de " + año_antier;
  //console.log(`Fecha antier es: ${antier_completo} \nFecha antier2: ${antier_completo2}`);

  //Se obtienen las 2 fechas de antiayer
   const antiayer = new Date(hoy);
  antiayer.setDate(hoy.getDate() - 3); // Restamos 1 día
  const año_antiayer = antiayer.getFullYear(); // Obtiene el año (ej: 2025)
  let mes_antiayer_temp = antiayer.getMonth() + 1; // Obtiene el mes (0-11, por eso le sumamos 1)
  const dia_antiayer = antiayer.getDate();
  const antiayer_completo2 = dia_antiayer + "/" + mes_antiayer_temp + "/" + año_antiayer;
  let mes_antiayer = regresames(mes_antiayer_temp);
  const antiayer_completo =dia_antiayer + " de " + mes_antiayer + " de " + año_antiayer;
  //console.log(`Fecha antiayer es: ${antiayer_completo} \nFecha antiayer2: ${antiayer_completo2}`);

  //Se obtienen las 2 fechas de antiantier
  const antiantier = new Date(hoy);
  antiantier.setDate(hoy.getDate() - 4); // Restamos 1 día
  const año_antiantier = antiantier.getFullYear(); // Obtiene el año (ej: 2025)
  let mes_antiantier_temp = antiantier.getMonth() + 1; // Obtiene el mes (0-11, por eso le sumamos 1)
  const dia_antiantier = antiantier.getDate();
  const antiantier_completo2 =dia_antiantier + "/" + mes_antiantier_temp + "/" + año_antiantier;
  let mes_antiantier = regresames(mes_antiantier_temp);
  const antiantier_completo = dia_antiantier + " de " + mes_antiantier + " de " + año_antiantier;
  //console.log(`Fecha antiantier es: ${antiantier_completo} \nFecha antiantier2: ${antiantier_completo2}`);

  //Se obtienen las 2 fechas de antiantier
  const fechita = new Date(hoy);
  fechita.setDate(hoy.getDate() - 5); // Restamos 1 día
  const año_fechita = fechita.getFullYear(); // Obtiene el año (ej: 2025)
  let mes_fechita_temp = fechita.getMonth() + 1; // Obtiene el mes (0-11, por eso le sumamos 1)
  const dia_fechita = fechita.getDate();
  const fechita_completo2 =dia_fechita + "/" + mes_fechita_temp + "/" + año_fechita;
  let mes_fechita = regresames(mes_fechita_temp);
  const fechita_completo = dia_fechita + " de " + mes_fechita + " de " + año_fechita;
  console.log(`Fecha aux es: ${fechita_completo} \nFecha aux2: ${fechita_completo2}`);

  //hoy_aux = antiayer_completo;

  //aqui empieza el ciclo 
  let bandera = true;
  let n = 1;
  let cuentan = 0;
  let cuenta = 0;
  const palabra = "Llantas";
  let numero = 0;
  let numero2 = 0;

 await  delay(10000);

  //Se obtiene el nuero de pedidos
  const xpath_numeropedidos = "/html/body/div[1]/div[2]/div/div/div[3]/div[4]/div[2]/div[2]/div[1]/div/span[1]";
                            
  var numero_pedidos = await obtenerTextoPorXPath(page, xpath_numeropedidos);

  console.log(`numero pedidos: s${numero_pedidos}`);
  if(!numero_pedidos)
  {
      const xpath_numeropedidos = '/html/body/div[1]/div[2]/div/div/div[3]/div[5]/div[2]/div[2]/div[1]/div/span[1]';
                            
      const numero_pedidos = await obtenerTextoPorXPath(page, xpath_numeropedidos);
  }
  if(!numero_pedidos)
  {
var holi ="/html/body/div[1]/div[2]/div/div/div[3]/div[5]/div[2]/div[2]/div[1]/div/span[1]";
          //*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]
         numero_pedidos = await page.evaluate((xpath) => {
          var el2 = document.evaluate(
            xpath,
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
          ).singleNodeValue;
          return el2 ? el2.textContent.trim() : null;
        }, holi); // aquí pasas la variable como
  }

  //Se divide el numero de pedidos para obtener el numero 

  const entero_numero_pedidos = await extraerNumerosDeString(numero_pedidos);
  console.log("Número de pedidos:", entero_numero_pedidos);
  
  //Aqui se manda a llamar la funcion con elnombre del archivo y se crea el excel
  const nombreArchivo = dia +"_" + mes_temp + "_" +año +"_" +hour +"_" +minute +"_" +milisec +".xlsx";
  crearArchivoExcel(nombreArchivo);
  const filePath = "./"+nombreArchivo;
  //Aqui comienza el ciclo :) 
  
  while (bandera != false)
  {
    //Todo esto para verificr si llegaa los 100 pedidos que de click en la nueva pagina mismo caso de 200 y 300 pedidos
    if (n > 100 && cuentan == 0) {
      
      const xpathBoton = '//*[@id="myo-layout"]/div[2]/div[4]/div/div/div[1]/div[2]/div/ul/li[2]/a';
      const clicExitoso = await clickElementoPorXPath(page, xpathBoton);
        if (clicExitoso) {
          console.log("¡Clic en el botón exitoso!");
        } else {
          console.log("No se pudo hacer clic en el botón.");
        }
      n = 1;
      cuenta = 100;
      cuentan = 1;
    } else if (n > 100 && cuentan == 1) {
      ////*[@id="myo-layout"]/div[2]/div[4]/div/div/div[1]/div[2]/div/ul/li[2]/a  /html/body/div[1]/div[2]/div/div/div[3]/div[4]/div[2]/div[5]/div/div/div[1]/div/div/ul/li[2]
      const xpathBoton = '//*[@id="myo-layout"]/div[2]/div[4]/div/div/div[1]/div[2]/div/ul/li[2]/a';
      const clicExitoso = await clickElementoPorXPath(page, xpathBoton);
        if (clicExitoso) {
          console.log("¡Clic en el botón exitoso!");
        } else {
          console.log("No se pudo hacer clic en el botón.");
        }
      n = 1;
      cuenta = 200;
      cuentan = 2;
    } else if (n > 100 && cuentan == 2) {
      const xpathBoton = '//*[@id="myo-layout"]/div[2]/div[4]/div/div/div[1]/div[2]/div/ul/li[2]/a';
      const clicExitoso = await clickElementoPorXPath(page, xpathBoton);
        if (clicExitoso) {
          console.log("¡Clic en el botón exitoso!");
        } else {
          console.log("No se pudo hacer clic en el botón.");
        }
      n = 1;
      cuenta = 300;
      cuentan = 3;
    }
    await delay(2000);
    //Se obtiene la fecha1 de los pedidos 
    let xpath_fechapedidos ="/html/body/div[1]/div[2]/div/div/div[3]/div[4]/div[2]/div[4]/div/div[2]/table/tbody/tr[" + n +"]/td[2]/div/div[2]/div";
    let fecha_pedidos1 = await obtenerTextoPorXPath(page, xpath_fechapedidos);

    if(!fecha_pedidos1)
    { 
                           
      var holi ="/html/body/div[1]/div[2]/div/div/div[3]/div[5]/div[2]/div[4]/div/div[2]/table/tbody/tr[" + n +"]/td[2]/div/div[2]/div";
          //*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]
         fecha_pedidos1 = await page.evaluate((xpath) => {
          var el2 = document.evaluate(
            xpath,
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
          ).singleNodeValue;
          return el2 ? el2.textContent.trim() : null;
        }, holi); // aquí pasas la variable como
    }
    console.log("Fecha de pedidos1:",fecha_pedidos1);
    //Se compara si la fecha es diferente de 4 dias si no se termina el script
    if (
      fecha_pedidos1 != hoy_completo2 &&
      fecha_pedidos1 != ayer_completo2 &&
      fecha_pedidos1 != antier_completo2 &&
      fecha_pedidos1 != antiayer_completo2 &&
      fecha_pedidos1 != antiantier_completo2 &&
      fecha_pedidos1 != fechita_completo2
    ) {
      bandera = false;
      break;
    }else{
     //Se clickea en el numero de pedido para la obtencion de los datos
     await delay(1000);
     let xpath_clickpedido = '//*[@id="orders-table"]/tbody/tr[' + n + "]/td[3]/div/div[1]/a";
     let clicExitoso = await clickElementoPorXPath(page, xpath_clickpedido);


      //Se espera a que cargeue la pagina 
      await delay(5000);

      let t = 0;
      let banderona = 0;
      
      //Se verifica que se tengan mas de una guia e edido aqui se pondrian esos casos
      try {
        let xpath_masdeunenvio ='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[2]/div[1]/div[2]/div/div[1]/div/div[2]';
                                
        let masdeunenvioElemento =  await obtenerTextoPorXPath(page, xpath_masdeunenvio);

        if (masdeunenvioElemento.length > 0) {
          t = 1;
          banderona = 1;
          console.log(`Hay mas de una guia en el pedido:)`)
        }
      } catch (error) {
        console.log("No está el elemento");
      } 

      //Aqui se verifica si se tienen dos pedidos diferentes en una misma guia 

      try {
        let xpath_masdeunenvio2 ='/html/body/div[1]/div[2]/div/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[2]/td[1]/div/div[1]/span/span/span';
        let masdeunenvioElemento2 =  await obtenerTextoPorXPath(page, xpath_masdeunenvio2);

        if (masdeunenvioElemento2.length > 0) {
          t = 1;
          banderona = 2;
          console.log(`Hay mas de una guia en el pedido:) de diferentes elementos`)
        }
      } catch (error) {
        console.log("No está el elemento");
      } 

      if (t == 1 && banderona == 1) {
      console.log(`Entro aqui alde 2 elementos y 2 numeor de guias distintos`);
              //aqui se comienza con el extraer datos la fecha del pedido
        let xpath_fechped='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[1]/div[1]/div[2]/div/div[1]/div/div[2]';
        let fechped = await obtenerTextoPorXPath(page, xpath_fechped);
       
        if (fechped == null) {
        let xpath_fechped= '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[1]/div[1]/div[2]/div/div[1]/div/div[2]';
        let fechped = await obtenerTextoPorXPath(page, xpath_fechped);
        }
          if (fechped == null) {
        let xpath_fechped= '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[1]/div[1]/div[2]/div/div[1]/div/div[2]';
        let fechped = await obtenerTextoPorXPath(page, xpath_fechped);
        }

        //fechped almacena la fecha
        fechped = fechped.slice(-17);
        console.log("fecha pedido en excel: " + fechped);

                //aqui comienza a extraer los datos del numped
        let xpath_numped='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[1]/div[1]/div/span[5]';
        let numped = await obtenerTextoPorXPath(page, xpath_numped);
        
        if (numped == null)  {
        let xpath_numped='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[1]/div[1]/div/span[5]';
        let numped = await obtenerTextoPorXPath(page, xpath_numped);
        }
        //se imprime el num de pedido
        console.log("Numpedido: " + numped);

        //Se empiza a extraer el numseg 
        let xpath_numseg='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[1]/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a';
        let numseg = await obtenerTextoPorXPath(page, xpath_numseg);
        
        if (numseg == null)  {
        let xpath_numseg='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[1]/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a';
        let numseg = await obtenerTextoPorXPath(page, xpath_numseg);
        }
        // almacena lel numero de guia o seguimiento
        console.log("Numseg: " + numseg);

                   //Se empieza a amacenar el nombre
        let xpath_nombre='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/table/tbody/tr[1]/td/span/span[3]/div/div/span[1]';
        let nombre = await obtenerTextoPorXPath(page, xpath_nombre);

        if (nombre == null) {
        let xpath_nombre='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/table/tbody/tr[1]/td/span/span[3]/div/div/span[1]';
        let nombre = await obtenerTextoPorXPath(page, xpath_nombre);
        }
        //almacena el nombre
        console.log("Nombre: " + nombre);

                //se empieza a almacenar el telefno 

        let xpath_telefono='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div/table/tbody/div/tr/td[2]/span';
        let telefono = await obtenerTextoPorXPath(page, xpath_telefono);

        if (telefono == null) {

        let xpath_telefono='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div/table/tbody/div/tr/td[2]/span';
        let telefono = await obtenerTextoPorXPath(page, xpath_telefono);
        }

        if (telefono == "") {
          telefono = 0;
        }
        //fechped almacena el telefono
        console.log("Telefono: " + telefono);

                //Aqui se almacena el titulo 1 
        let xpath_llantas='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[1]/div[2]/div/table/tbody/tr/td[2]/div/div[1]/div/a/div';
                          '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[2]/div[2]/div/table/tbody/tr/td[2]/div/div[1]/div/a/div';
                          
        let llantas = await obtenerTextoPorXPath(page, xpath_llantas);
        if (llantas == null)  {
        let xpath_llantas='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[1]/div[2]/div/table/tbody/tr/td[2]/div/div[1]/div/a/div';
        let llantas = await obtenerTextoPorXPath(page, xpath_llantas);
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Titulo: " + llantas);

        let xpath_llantas2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[2]/div[2]/div/table/tbody/tr/td[2]/div/div[1]/div/a/div';                
        let llantas2 = await obtenerTextoPorXPath(page, xpath_llantas2);
        if (llantas2 == null)  {
        let xpath_llantas2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr[2]/td[3]/div/div[1]/div/a/div';
        let llantas2 = await obtenerTextoPorXPath(page, xpath_llantas2);
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Titulo: " + llantas2);

        await delay(800);

                //Se comienza a extraer la cantidad 
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[1]/div[2]/div/table/tbody/tr/td[4]';;
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
          
        if (cantidad == null)  {
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
        }
        if (cantidad == null)  {
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[2]/div/table/tbody/tr/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
        }
        if (cantidad == null)  {
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Cantidad: " + cantidad);

        let xpath_cantidad2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[2]/div[2]/div/table/tbody/tr/td[4]';
        '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[2]/div[2]/div/table/tbody/tr/td[4]';
        let cantidad2 = await obtenerTextoPorXPath(page, xpath_cantidad2);
          
        if (cantidad2 == null)  {
        let xpath_cantidad2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[5]';
        let cantidad2 = await obtenerTextoPorXPath(page, xpath_cantidad2);
        }
        if (cantidad2 == null)  {
        let xpath_cantidad2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[2]/div/table/tbody/tr/td[5]';
        let cantidad2 = await obtenerTextoPorXPath(page, xpath_cantidad2);
        }
        if (cantidad2 == null)  {
        let xpath_cantidad2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[5]';
        let cantidad2 = await obtenerTextoPorXPath(page, xpath_cantidad2);
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Cantidad 2: " + cantidad2);

                  //Se comienza a extraer el dato del precio 
        let xpath_precio='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[1]/div[2]/div/table/tbody/tr/td[5]/div/table[1]/tbody/div[3]/div[2]/span';
        let precio = await obtenerTextoPorXPath(page, xpath_precio);
        if (precio == null) {
        let xpath_precio='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr[1]/td[6]/span';
        let precio = await obtenerTextoPorXPath(page, xpath_precio);
        }if (precio == null) {
        let xpath_precio= '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[1]/td[7]/span';
        let precio = await obtenerTextoPorXPath(page, xpath_precio);
        }

       
        //fechped almacena el precio
        console.log("Precio: " + precio);

                //Se comienza a extraer el dato del precio 
        let xpath_precio2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[6]/span';
        let precio2 = await obtenerTextoPorXPath(page, xpath_precio2);
        if (precio2 == null) {
        let xpath_precio2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr[2]/td[6]/span';
        let precio2 = await obtenerTextoPorXPath(page, xpath_precio2);
        }if (precio2 == null) {
        let xpath_precio2= '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[2]/td[7]/span';
        let precio2 = await obtenerTextoPorXPath(page, xpath_precio2);
        }

       
        //fechped almacena el precio
        console.log("Precio2: " + precio2);

        hoy_aux = hoy_aux.trim().normalize();
        fechped = fechped.trim().normalize();
        
        console.log(`Hoy aux: ${hoy_aux}\nfechped: ${fechped}`);
        if (fechped == hoy_aux)
        {
        if (llantas.includes(palabra)) {
          for (let i = 0; i < llantas.length; i++) {
            if (/\d/.test(llantas[i])) {
              numero = parseInt(llantas[i]);
              break;
            }
          }
        } else {
          numero = 1;
        }

        if (llantas2.includes(palabra)) {
          for (let i = 0; i < llantas2.length; i++) {
            if (/\d/.test(llantas2[i])) {
              numero2 = parseInt(llantas2[i]);
              break;
            }
          }
        } else {
          numero2 = 1;
        }        
      
      console.log(`llantaspor paq= ${numero}`);
      let preciolimpio = precio.replace("$", "");
      preciolimpio = preciolimpio.replace(",", "");
      console.log(` ${preciolimpio}`);
      let precioEntero = parseInt(preciolimpio);
      console.log(` ${precioEntero}`);
      let ventatot = precioEntero * parseFloat(cantidad);
      
      console.log(` ${ventatot}`);
      let pusi = (ventatot / parseFloat(numero))/1.16; 
      console.log(` ${pusi}`);
      let llantash=llantas;
      //console.log(`holiis`);
      console.log(`${llantas} ${cantidad}`);
      let llantastot= parseFloat(numero)*parseFloat(cantidad);
      console.log(llantastot);

            console.log(`llantaspor paq= ${numero}`);
      let preciolimpio2 = precio2.replace("$", "");
      preciolimpio2 = preciolimpio2.replace(",", "");
      console.log(` ${preciolimpio2}`);
      let precioEntero2 = parseInt(preciolimpio2);
      console.log(` ${precioEntero2}`);
      let ventatot2 = precioEntero2 * parseFloat(cantidad2);
      
      console.log(` ${ventatot2}`);
      let pusi2 = (ventatot2 / parseFloat(numero2))/1.16; 
      console.log(` ${pusi2}`);
      let llantash2=llantas2;
      //console.log(`holiis`);
      console.log(`${llantas2} ${cantidad2}`);
      let llantastot2= parseFloat(numero2)*parseFloat(cantidad2);
      console.log(llantastot);

      const pedidoEjemplo2 = [numped,llantastot,llantas,cantidad,precio,'',ventatot,pusi,'',numseg,nombre,telefono,'','',''];
      let numerodecelda = cuenta+ n + 1;
      let hoja=1;
      await delay(500); // Agregar un delay antes de editar
      await editarExcel(filePath,numerodecelda,pedidoEjemplo2,hoja);

      await delay(500); // Agregar un delay antes de editar
      const pedidoEjemplo3 = [numped,llantastot2,llantas2,cantidad2,precio2,'',ventatot2,pusi2,'',numseg,nombre,telefono,'','',''];
      let numerodecelda2 = cuenta+ n + 1;
      let hoja2=1;
      await editarExcel(filePath,numerodecelda2,pedidoEjemplo3,hoja2);
      }






      
      n=n+1;
      }else if (t==1 && banderona == 2)
      {
         console.log(`Entro aqui alde 1 guia y 2 aquetes diferentes`);

        //aqui se comienza con el extraer datos la fecha del pedido
        let xpath_fechped='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[1]/div/div[2]';
        let fechped = await obtenerTextoPorXPath(page, xpath_fechped);
       
        if (fechped == null) {
        let xpath_fechped= '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]';
        let fechped = await obtenerTextoPorXPath(page, xpath_fechped);
        }
          if (fechped == null) {
        let xpath_fechped= '/html/body/div[1]/div[2]/div/div/div[1]/div[1]/div/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div[2]';
        let fechped = await obtenerTextoPorXPath(page, xpath_fechped);
        }

        //fechped almacena la fecha
        fechped = fechped.slice(-17);
        console.log("fecha pedido en excel: " + fechped);

        //aqui comienza a extraer los datos del numped
        let xpath_numped='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[1]/div[1]/div/span[5]';
        let numped = await obtenerTextoPorXPath(page, xpath_numped);
        
        if (numped == null)  {
        let xpath_numped='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[1]/div[1]/div/span[5]';
        let numped = await obtenerTextoPorXPath(page, xpath_numped);
        }
        //se imprime el num de pedido
        console.log("Numpedido: " + numped);

        //Se empiza a extraer el numseg 
        let xpath_numseg='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a';
        let numseg = await obtenerTextoPorXPath(page, xpath_numseg);
        
        if (numseg == null)  {
        let xpath_numseg='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a';
        let numseg = await obtenerTextoPorXPath(page, xpath_numseg);
        }
        // almacena lel numero de guia o seguimiento
        console.log("Numseg: " + numseg);

            //Se empieza a amacenar el nombre
        let xpath_nombre='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/table/tbody/tr[1]/td/span/span[3]/div/div/span[1]';
        let nombre = await obtenerTextoPorXPath(page, xpath_nombre);

        if (nombre == null) {
        let xpath_nombre='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/table/tbody/tr[1]/td/span/span[3]/div/div/span[1]';
        let nombre = await obtenerTextoPorXPath(page, xpath_nombre);
        }
        //almacena el nombre
        console.log("Nombre: " + nombre);
        //

        //se empieza a almacenar el telefno 

        let xpath_telefono='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div/table/tbody/div/tr/td[2]/span';
        let telefono = await obtenerTextoPorXPath(page, xpath_telefono);

        if (telefono == null) {

        let xpath_telefono='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div/table/tbody/div/tr/td[2]/span';
        let telefono = await obtenerTextoPorXPath(page, xpath_telefono);
        }

        if (telefono == "") {
          telefono = 0;
        }
        //fechped almacena el telefono
        console.log("Telefono: " + telefono);
        //Aqui se almacena el titulo 1 
        let xpath_llantas='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[1]/td[3]/div/div[1]/div/a/div';
                          
        let llantas = await obtenerTextoPorXPath(page, xpath_llantas);
        if (llantas == null)  {
        let xpath_llantas='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr[1]/td[3]/div/div[1]/div/a/div';
        let llantas = await obtenerTextoPorXPath(page, xpath_llantas);
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Titulo: " + llantas);

        let xpath_llantas2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[2]/td[3]/div/div[1]/div/a/div';                  
        let llantas2 = await obtenerTextoPorXPath(page, xpath_llantas2);
        if (llantas2 == null)  {
        let xpath_llantas2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr[2]/td[3]/div/div[1]/div/a/div';
        let llantas2 = await obtenerTextoPorXPath(page, xpath_llantas2);
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Titulo: " + llantas2);

        await delay(800);
        //Se comienza a extraer la cantidad 
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[1]/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
          
        if (cantidad == null)  {
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
        }
        if (cantidad == null)  {
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[2]/div/table/tbody/tr/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
        }
        if (cantidad == null)  {
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Cantidad: " + cantidad);



        let xpath_cantidad2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[2]/td[5]';
        '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div[2]/div[2]/div/table/tbody/tr/td[4]';
        let cantidad2 = await obtenerTextoPorXPath(page, xpath_cantidad2);
          
        if (cantidad2 == null)  {
        let xpath_cantidad2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[5]';
        let cantidad2 = await obtenerTextoPorXPath(page, xpath_cantidad2);
        }
        if (cantidad2 == null)  {
        let xpath_cantidad2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[2]/div/table/tbody/tr/td[5]';
        let cantidad2 = await obtenerTextoPorXPath(page, xpath_cantidad2);
        }
        if (cantidad2 == null)  {
        let xpath_cantidad2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[5]';
        let cantidad2 = await obtenerTextoPorXPath(page, xpath_cantidad2);
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Cantidad 2: " + cantidad2);


                  //Se comienza a extraer el dato del precio 
        let xpath_precio='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[1]/td[6]/span';
        let precio = await obtenerTextoPorXPath(page, xpath_precio);
        if (precio == null) {
        let xpath_precio='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr[1]/td[6]/span';
        let precio = await obtenerTextoPorXPath(page, xpath_precio);
        }if (precio == null) {
        let xpath_precio= '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[1]/td[7]/span';
        let precio = await obtenerTextoPorXPath(page, xpath_precio);
        }

       
        //fechped almacena el precio
        console.log("Precio: " + precio);


        //Se comienza a extraer el dato del precio 
        let xpath_precio2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[2]/td[6]/span';
        let precio2 = await obtenerTextoPorXPath(page, xpath_precio2);
        if (precio2 == null) {
        let xpath_precio2='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr[2]/td[6]/span';
        let precio2 = await obtenerTextoPorXPath(page, xpath_precio2);
        }if (precio2 == null) {
        let xpath_precio2= '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr[2]/td[7]/span';
        let precio2 = await obtenerTextoPorXPath(page, xpath_precio2);
        }

       
        //fechped almacena el precio
        console.log("Precio2: " + precio2);

        hoy_aux = hoy_aux.trim().normalize();
        fechped = fechped.trim().normalize();
        
        console.log(`Hoy aux: ${hoy_aux}\nfechped: ${fechped}`);
        if (fechped == hoy_aux)
        {
        if (llantas.includes(palabra)) {
          for (let i = 0; i < llantas.length; i++) {
            if (/\d/.test(llantas[i])) {
              numero = parseInt(llantas[i]);
              break;
            }
          }
        } else {
          numero = 1;
        }

        if (llantas2.includes(palabra)) {
          for (let i = 0; i < llantas2.length; i++) {
            if (/\d/.test(llantas2[i])) {
              numero2 = parseInt(llantas2[i]);
              break;
            }
          }
        } else {
          numero2 = 1;
        }        
      
      console.log(`llantaspor paq= ${numero}`);
      let preciolimpio = precio.replace("$", "");
      preciolimpio = preciolimpio.replace(",", "");
      console.log(` ${preciolimpio}`);
      let precioEntero = parseInt(preciolimpio);
      console.log(` ${precioEntero}`);
      let ventatot = precioEntero * parseFloat(cantidad);
      
      console.log(` ${ventatot}`);
      let pusi = (ventatot / parseFloat(numero))/1.16; 
      console.log(` ${pusi}`);
      let llantash=llantas;
      //console.log(`holiis`);
      console.log(`${llantas} ${cantidad}`);
      let llantastot= parseFloat(numero)*parseFloat(cantidad);
      console.log(llantastot);

            console.log(`llantaspor paq= ${numero}`);
      let preciolimpio2 = precio2.replace("$", "");
      preciolimpio2 = preciolimpio2.replace(",", "");
      console.log(` ${preciolimpio2}`);
      let precioEntero2 = parseInt(preciolimpio2);
      console.log(` ${precioEntero2}`);
      let ventatot2 = precioEntero2 * parseFloat(cantidad2);
      
      console.log(` ${ventatot2}`);
      let pusi2 = (ventatot2 / parseFloat(numero2))/1.16; 
      console.log(` ${pusi2}`);
      let llantash2=llantas2;
      //console.log(`holiis`);
      console.log(`${llantas2} ${cantidad2}`);
      let llantastot2= parseFloat(numero2)*parseFloat(cantidad2);
      console.log(llantastot);

      const pedidoEjemplo2 = [numped,llantastot,llantas,cantidad,precio,'',ventatot,pusi,'',numseg,nombre,telefono,'','',''];
      let numerodecelda = cuenta+ n + 1;
      let hoja=1;
      await delay(500); // Agregar un delay antes de editar
      await editarExcel(filePath,numerodecelda,pedidoEjemplo2,hoja);

      await delay(500); // Agregar un delay antes de editar
      const pedidoEjemplo3 = [numped,llantastot2,llantas2,cantidad2,precio2,'',ventatot2,pusi2,'',numseg,nombre,telefono,'','',''];
      let numerodecelda2 = cuenta+ n + 1;
      let hoja2=1;
      await editarExcel(filePath,numerodecelda2,pedidoEjemplo3,hoja2);
        }

        console.log(`\n`);



        n=n+1;
      }else
      {
      
         await delay(2000);
        //aqui se comienza con el extraer datos la fecha del pedido
        console.log(`aqui se entra a la 3 opcion cuando solo es una guia y un pedido`);
       
        let xpath_fechped='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[1]/div/div[2]';
        let fechped = await obtenerTextoPorXPath(page, xpath_fechped);
       
        if (fechped == null) {
          console.log(`2`);
        let xpath_fechped= '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]';
                            //*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]
        let fechped = await obtenerTextoPorXPath(page, xpath_fechped);
        }
          if (fechped == null) {
            console.log(`3`);
        let xpath_fechped= '/html/body/div[1]/div[2]/div/div/div[1]/div[1]/div/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div[2]';
        let fechped = await obtenerTextoPorXPath(page, xpath_fechped);
        }
        if ( fechped== null)
        {
         
        var holi ='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]';
          //*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]
         fechped = await page.evaluate((xpath) => {
          var el2 = document.evaluate(
            xpath,
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
          ).singleNodeValue;
          return el2 ? el2.textContent.trim() : null;
        }, holi); // aquí pasas la variable como
       
        }
        console.log("fecha pedido en excel: " + fechped);
        //fechped almacena la fecha
        fechped = fechped.slice(-17);
        


        //aqui comienza a extraer los datos del numped

        let xpath_numped='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[1]/div[1]/div/span[5]';
        let numped = await obtenerTextoPorXPath(page, xpath_numped);
        
        if (numped == null)  {
        let xpath_numped='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[1]/div[1]/div/span[5]';
        let numped = await obtenerTextoPorXPath(page, xpath_numped);
        }
        //se imprime el num de pedido
        console.log("Numpedido: " + numped);


        //Se empiza a extraer el numseg 
        let xpath_numseg='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a';
        let numseg = await obtenerTextoPorXPath(page, xpath_numseg);
        
        if (numseg == null)  {
        let xpath_numseg='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a';
        let numseg = await obtenerTextoPorXPath(page, xpath_numseg);
        }

        if (numseg==null)
        {
          var holi ='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a/text()';
          //*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]
         numseg = await page.evaluate((xpath) => {
          var el2 = document.evaluate(
            xpath,
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
          ).singleNodeValue;
          return el2 ? el2.textContent.trim() : null;
        }, holi); // aquí pasas la variable como
       
        }
        // almacena lel numero de guia o seguimiento
        console.log("Numseg: " + numseg);
        //Se empieza a amacenar el nombre
        let xpath_nombre='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/table/tbody/tr[1]/td/span/span[3]/div/div/span[1]';
        let nombre = await obtenerTextoPorXPath(page, xpath_nombre);

        if (nombre == null) {
        let xpath_nombre='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/table/tbody/tr[1]/td/span/span[3]/div/div/span[1]';
        let nombre = await obtenerTextoPorXPath(page, xpath_nombre);
        }
        //almacena el nombre
        console.log("Nombre: " + nombre);
        //

        //se empieza a almacenar el telefno 

        let xpath_telefono='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div/table/tbody/div/tr/td[2]/span';
        let telefono = await obtenerTextoPorXPath(page, xpath_telefono);

        if (telefono == null) {

        let xpath_telefono='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div/table/tbody/div/tr/td[2]/span';
        let telefono = await obtenerTextoPorXPath(page, xpath_telefono);
        }

        if (telefono == "") {
          telefono = 0;
        }
        //fechped almacena el telefono
        console.log("Telefono: " + telefono);

        //titulo de las llantes

        let xpath_llantas='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[3]/div/div[1]/div/a/div';
        let llantas = await obtenerTextoPorXPath(page, xpath_llantas);
        if (llantas == null)  {
        let xpath_llantas='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[3]/div/div[1]/div/a/div';
        let llantas = await obtenerTextoPorXPath(page, xpath_llantas);
        }

        if(llantas == null)
        {
          var holi ='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[3]/div/div[1]/div/a/div';
                     //*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[3]/div/div[1]/div/a/div
          //*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]
        llantas = await page.evaluate((xpath) => {
          var el2 = document.evaluate(
            xpath,
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
          ).singleNodeValue;
          return el2 ? el2.textContent.trim() : null;
        }, holi); // aquí pasas la variable como
       
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Titulo: " + llantas);
        
        await delay(800);
        //Se comienza a extraer la cantidad 
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[2]/div/table/tbody/tr/td[4]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
          
        if (cantidad == null)  {
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
        }
        if (cantidad == null)  {
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[2]/div/table/tbody/tr/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
        }
        if (cantidad == null)  {
        let xpath_cantidad='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[5]';
        let cantidad = await obtenerTextoPorXPath(page, xpath_cantidad);
        }

        if(cantidad == null)
        {
          var holi ='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[5]';
          //*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]
         cantidad = await page.evaluate((xpath) => {
          var el2 = document.evaluate(
            xpath,
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
          ).singleNodeValue;
          return el2 ? el2.textContent.trim() : null;
        }, holi); // aquí pasas la variable como
       
        }
        //fechped almacena el titulo para el numero de llantas
        console.log("Cantidad: " + cantidad);

          //Se comienza a extraer el dato del precio 
        let xpath_precio='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[6]/span';
        let precio = await obtenerTextoPorXPath(page, xpath_precio);
        if (precio == null) {
        let xpath_precio='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[6]/span';
        let precio = await obtenerTextoPorXPath(page, xpath_precio);
        }if (precio == null) {
        let xpath_precio= '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[7]/span';
        let precio = await obtenerTextoPorXPath(page, xpath_precio);
        }
        
        if(precio == null)
        {
          var holi ='//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[6]/span';
          //*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[1]/div/div[2]
         precio = await page.evaluate((xpath) => {
          var el2 = document.evaluate(
            xpath,
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
          ).singleNodeValue;
          return el2 ? el2.textContent.trim() : null;
        }, holi); // aquí pasas la variable como
       
        }
       
        //fechped almacena el precio
        console.log("Precio: " + precio);
       
        console.log(`Hoy completo: ${hoy_aux}`);

        hoy_aux = hoy_aux.trim().normalize();
        fechped = fechped.trim().normalize();
        
        if (fechped == hoy_aux)
        {
        if (llantas.includes(palabra)) {
          for (let i = 0; i < llantas.length; i++) {
            if (/\d/.test(llantas[i])) {
              numero = parseInt(llantas[i]);
              break;
            }
          }
        } else {
          numero = 1;
        }
      
      //  console.log(`llantaspor paq= ${numero}`);
      let preciolimpio = precio.replace("$", "");
      preciolimpio = preciolimpio.replace(",", "");
      //console.log(` ${preciolimpio}`);
      let precioEntero = parseInt(preciolimpio);
      //console.log(` ${precioEntero}`);
      let ventatot = precioEntero * parseFloat(cantidad);
      
      //console.log(` ${ventatot}`);
      let pusi = (ventatot / parseFloat(numero))/1.16; 
      //console.log(` ${pusi}`);
      let llantash=llantas;
      //console.log(`holiis`);
      //console.log(`${llantas} ${cantidad}`);
      let llantastot= parseFloat(numero)*parseFloat(cantidad);
      //console.log(llantastot);
      const pedidoEjemplo = [numped,llantastot,llantas,cantidad,precio,'',ventatot,pusi,'',numseg,nombre,telefono,'','',''];
      let numerodecelda = cuenta+ n + 1;
      let hoja=1;
      await delay(500); // Agregar un delay antes de editar
      await editarExcel(filePath,numerodecelda,pedidoEjemplo,hoja);



        }
      }
      
      

    }






        await page.goto(
        "https://sellercentral.amazon.com.mx/orders-v3/mfn/shipped/selfship?ref_=xx_swlang_head_xx&mons_sel_locale=es_MX&languageSwitched=1&page=1",
        { waitUntil: "domcontentloaded" }
        );
        n = n + 1;
        await delay(5000);
        console.log(`\n`);

  }


  console.log(`El Sript se ha terminado se obtuvieron los datos del diacde hoy en un periodo de 5 dias :) .`);
}



run();
