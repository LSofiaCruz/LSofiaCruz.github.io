
const excelFilePath = '../assets/data.xlsx';
let allDataBK = {}

const processListByChar = function(line, splitChar){
    const cleanLine = line.split(splitChar).filter(item => item.trim() !== '')
    let cleanedDiv = '<div><ul>'

    cleanLine.forEach(item => {
        cleanedDiv += `
            <li>${item}</li>
        `
    })

    cleanedDiv += '</ul></div>'
    return cleanedDiv
}

const renderEXCEL = async function(allJSONData){

    const outputDiv = document.getElementById('output');
    allJSONData.forEach(row => {
        const nombre_ing = row["INGREDIENTE"] || '';
        const compuestos_ing = row["COMPUESTOS ACTIVOS"] || '';
        const biodisponibilidad_ing = row["BIODISPONIBILIDAD"] || '';
        const literatura_ing = row["QUE DICE LA LITERATURA"] || '';
        const tipos_fuentes_ing = row["TIPOS DE FUENTES"] || '';
        const matrices_ing = row["TIPO DE ALIMENTOS O MATRICES"] || '';
        const formas_ing = row["FORMAS"] || '';
        const most_used_ing = row["FORMA MÁS UTILIZADA"] || ''
        const bio_factor_ing = row["FACTORES DE BIODISPONIBILIDAD"] || '';
        const dosis_ing = row["DOSIS"] || '';
        const ul_ing = row["UL"] || '';
        const used_feed_ing = row["FORMAS USADAS EN ALIMENTOS"] || '';
        const ref_ing = row["REFERENCIAS"] || '';

        const clean_bio = processListByChar(biodisponibilidad_ing, "• ");
        const clean_lit = processListByChar(literatura_ing, "• ");
        const clean_font = processListByChar(tipos_fuentes_ing, "• ");
        const clean_mat = processListByChar(matrices_ing, "• ");
        const clean_form = processListByChar(formas_ing, "• ");
        const clean_most = processListByChar(most_used_ing, "• ");
        const clean_bio_fact = processListByChar(bio_factor_ing, "• ");
        const clean_dosis = processListByChar(dosis_ing, "• ");
        const clean_ul = processListByChar(ul_ing, "• ");
        const clean_used_food = processListByChar(used_feed_ing, "• ");
        const clean_ref = processListByChar(ref_ing, "• ");

        const section = document.createElement('div');
        section.innerHTML = `
            <div class="card">
                <div class="card-header">
                    <h1>${nombre_ing} <span></span></h1>
                </div>
                <div class="card-body">
                    <div class="${compuestos_ing.trim() == ''? "hidden":"display"}">
                        <h4>COMPUESTOS ACTIVOS:</h4> 
                        <span>${compuestos_ing}</span>
                    </div>
                    <div class="${biodisponibilidad_ing.trim() == ''? "hidden":"display"}">
                        <h4>BIODISPONIBILIDAD:</h4>
                        <span>${clean_bio}</span>
                        <div class="${bio_factor_ing.trim() == ''? "hidden":"display"}">
                            <h4>FACTORES DE BIODISPONIBILIDAD:</h4> 
                            <span>${clean_bio_fact}</span>
                        </div>
                    </div>
                    <div class="${literatura_ing.trim() == ''? "hidden":"display"}">
                        <h4>QUE DICE LA LITERATURA:</h4>
                        <span>${clean_lit}</span>
                    </div>
                    <div class="${tipos_fuentes_ing.trim() == ''? "hidden":"display"}">
                        <h4>TIPOS DE FUENTES:</h4>
                        <span>${clean_font} </span>
                    </div>
                    <div class="${matrices_ing.trim() == ''? "hidden":"display"}">
                        <h4>TIPO DE ALIMENTOS O MATRICES:</h4>
                        <span>${clean_mat} </span>
                    </div>
                    <div class="${formas_ing.trim() == ''? "hidden":"display"} sub-card">
                        <h4>FORMAS:</h4>
                        <span>${clean_form} </span>
                        <div class="${most_used_ing.trim() == ''? "hidden":"display"} sub-card-section">
                            <h5>LAS MÁS USADAS EN GENERAL:</h4> 
                            <span>${clean_most} </span>
                        </div>
                        <div class="${used_feed_ing.trim() == ''? "hidden":"display"} sub-card-section">
                            <h5>LAS MÁS USADAS EN ALIMENTOS:</h4> 
                            <span>${clean_used_food} </span>
                        </div>
                    </div>
                    <div class="${dosis_ing.trim() == ''? "hidden":"display"}">
                        <h4>DOSIS:</h4> 
                        <span>${clean_dosis} </span>
                    </div>
                    <div class="${ul_ing.trim() == ''? "hidden":"display"}">
                        <h4>UL:</h4>
                        <span>${clean_ul} </span>
                    </div>
                </div>
                <div class="card-footer">
                    <p>${clean_ref}</p>
                </div>
            </div>
        `
        outputDiv.appendChild(section);
    });
}

const preProcessSheets = function(excelWorkbook) {
    let finalJSON = []
    excelWorkbook.SheetNames.forEach(hoja => {
        const miniJSON = XLSX.utils.sheet_to_json(excelWorkbook.Sheets[hoja]);
        finalJSON = finalJSON.concat(miniJSON);
    })
    return finalJSON;
}

fetch(excelFilePath)
    .then(response => response.arrayBuffer())
    .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const jsonData = preProcessSheets(workbook)
        allDataBK = jsonData;
        renderEXCEL(jsonData);
    })
    .catch(error => {
        console.error('Error al leer el archivo Excel:', error);
    });
