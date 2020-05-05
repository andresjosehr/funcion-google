/*************************************************************/
// NumeroALetras
// The MIT License (MIT)
// 
// Copyright (c) 2015 Luis Alfredo Chee 
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
// 
// @author Rodolfo Carmona
// @contributor Jean (jpbadoino@gmail.com)
/*************************************************************/
function Unidades(num){

    switch(num)
    {
        case 1: return "UN";
        case 2: return "DOS";
        case 3: return "TRES";
        case 4: return "CUATRO";
        case 5: return "CINCO";
        case 6: return "SEIS";
        case 7: return "SIETE";
        case 8: return "OCHO";
        case 9: return "NUEVE";
    }

    return "";
}//Unidades()

function Decenas(num){

    decena = Math.floor(num/10);
    unidad = num - (decena * 10);

    switch(decena)
    {
        case 1:
            switch(unidad)
            {
                case 0: return "DIEZ";
                case 1: return "ONCE";
                case 2: return "DOCE";
                case 3: return "TRECE";
                case 4: return "CATORCE";
                case 5: return "QUINCE";
                default: return "DIECI" + Unidades(unidad);
            }
        case 2:
            switch(unidad)
            {
                case 0: return "VEINTE";
                default: return "VEINTI" + Unidades(unidad);
            }
        case 3: return DecenasY("TREINTA", unidad);
        case 4: return DecenasY("CUARENTA", unidad);
        case 5: return DecenasY("CINCUENTA", unidad);
        case 6: return DecenasY("SESENTA", unidad);
        case 7: return DecenasY("SETENTA", unidad);
        case 8: return DecenasY("OCHENTA", unidad);
        case 9: return DecenasY("NOVENTA", unidad);
        case 0: return Unidades(unidad);
    }
}//Unidades()

function DecenasY(strSin, numUnidades) {
    if (numUnidades > 0)
    return strSin + " Y " + Unidades(numUnidades)

    return strSin;
}//DecenasY()

function Centenas(num) {
    centenas = Math.floor(num / 100);
    decenas = num - (centenas * 100);

    switch(centenas)
    {
        case 1:
            if (decenas > 0)
                return "CIENTO " + Decenas(decenas);
            return "CIEN";
        case 2: return "DOSCIENTOS " + Decenas(decenas);
        case 3: return "TRESCIENTOS " + Decenas(decenas);
        case 4: return "CUATROCIENTOS " + Decenas(decenas);
        case 5: return "QUINIENTOS " + Decenas(decenas);
        case 6: return "SEISCIENTOS " + Decenas(decenas);
        case 7: return "SETECIENTOS " + Decenas(decenas);
        case 8: return "OCHOCIENTOS " + Decenas(decenas);
        case 9: return "NOVECIENTOS " + Decenas(decenas);
    }

    return Decenas(decenas);
}//Centenas()

function Seccion(num, divisor, strSingular, strPlural) {
    cientos = Math.floor(num / divisor)
    resto = num - (cientos * divisor)

    letras = "";

    if (cientos > 0)
        if (cientos > 1)
            letras = Centenas(cientos) + " " + strPlural;
        else
            letras = strSingular;

    if (resto > 0)
        letras += "";

    return letras;
}//Seccion()

function Miles(num) {
    divisor = 1000;
    cientos = Math.floor(num / divisor)
    resto = num - (cientos * divisor)

    strMiles = Seccion(num, divisor, "UN MIL", "MIL");
    strCentenas = Centenas(resto);

    if(strMiles == "")
        return strCentenas;

    return strMiles + " " + strCentenas;
}//Miles()

function Millones(num) {
    divisor = 1000000;
    cientos = Math.floor(num / divisor)
    resto = num - (cientos * divisor)

    strMillones = Seccion(num, divisor, "UN MILLON", "MILLONES");
    strMiles = Miles(resto);

    if(strMillones == "")
        return strMiles;

    return strMillones + " " + strMiles;
}//Millones()

function NumeroALetras(num) {
    var data = {
        numero: num,
        enteros: Math.floor(num),
        centavos: (((Math.round(num * 100)) - (Math.floor(num) * 100))),
    };

    if (data.centavos > 0) {
        data.letrasCentavos = "CON " + (function (){
            if (data.centavos == 1)
                return Millones(data.centavos);
            else
                return Millones(data.centavos);
            })();
    };

    if(data.enteros == 0)
        return "CERO " + data.letrasMonedaPlural;
    if (data.enteros == 1)
        return Millones(data.enteros);
    else
        return Millones(data.enteros);
}//NumeroALetras()






/** @OnlyCurrentDoc */
function EmailCongelarDeuda() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(libro.getSheetByName("Propuestas deuda congelada"));
  const hoja = SpreadsheetApp.getActiveSheet();
  const filas = hoja.getRange("A2:S").getValues();

 
  for (indiceFila in filas) {
    var candidato = crearCandidato(filas[indiceFila]);
    enviarCorreo(candidato);
 }
}

function crearCandidato(datosFila) {
  const candidato = {
	númerounico: datosFila[0],
	propuestas: datosFila[1],
	fechadelarchivo: datosFila[2],
	nombre: datosFila[3],
	cedula: datosFila[4],
	email: datosFila[5],
	fechadeinicio: datosFila[6],
	montocuota: datosFila[7],
	montocuotaletras: NumeroALetras(FuncionQueQuitaSignoYComa(datosFila[7])),
	cupo: datosFila[9],
	contrato: datosFila[10],
	pagoinicial: datosFila[11],
	letrapagoinicial: NumeroALetras(FuncionQueQuitaSignoYComa(datosFila[11])),
	plazototal: datosFila[13],
	segundascuotas: datosFila[14],
	saldototal: datosFila[15],
	letrassaldototal: NumeroALetras(FuncionQueQuitaSignoYComa(datosFila[15])),
	totdescuento: datosFila[17],
	totaldescuentoletras:  NumeroALetras(FuncionQueQuitaSignoYComa(datosFila[17])),
	descuento: datosFila[19],
	pagocondescuento: datosFila[20],
	pagodescuentoletras: NumeroALetras(FuncionQueQuitaSignoYComa(datosFila[20])),
};
 
  return candidato;
}

function enviarCorreo(candidato) {
  if (candidato.email == "") return;
  const plantilla = HtmlService.createTemplateFromFile('ConDeu');
  plantilla.candidato = candidato;
  const mensaje = plantilla.evaluate().getContent();
 
  MailApp.sendEmail({
    to: candidato.email,
    subject: "Propuesta deuda congelada contrato " +candidato.contrato+ " y cupo " +candidato.cupo,
    htmlBody: mensaje,
   
  });
}

function FuncionQueQuitaSignoYComa(num) {
    //Proceso que quita la broma

        return num.replace(/[^0-9]/gi,'');

  
}

// console.log(FuncionQueQuitaSignoYComa("$25,249"));