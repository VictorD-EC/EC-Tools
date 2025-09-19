import { TabRow } from "./TabRow.js";

const formTabs = document.getElementById('formTabs')

const userFile = document.getElementById('userFile');
const sheetnameSelect = document.getElementById('sheetnameSelect')
const firstRow = document.getElementById('firstRow')
const descriptionInput = document.getElementById('inputDescription')
const uniteInput = document.getElementById('inputUnite')
const quantiteInput = document.getElementById('inputQuantite')
const totalInput = document.getElementById('inputTotal')
const readXlsxBtn = document.getElementById('readXlsx');
const generateXlsxBtn = document.getElementById('generateXlsx');

const userFile2 = document.getElementById('userFile2');
const readXlsxBtn2 = document.getElementById('readXlsx2');


const lotSelect = document.getElementById('lotSelect');
const sousLotSelect = document.getElementById('sousLotSelect');
const tbody = document.getElementById('tbody')
const mergedTableContainer = document.getElementById('mergedTable');

let workbook;

// affichage des noms des feuilles dans le select au moment ou un fichier est téléchargé
userFile.addEventListener("change", (e) => {
    const data = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (event) => {
        const binaryString = event.target.result;
        workbook = XLSX.read(binaryString, { type: 'binary' });
        sheetnameSelect.innerHTML = "";
        workbook.SheetNames.forEach(element => {
            let option = document.createElement("option");
            option.value = element;
            option.textContent = element;
            sheetnameSelect.appendChild(option);
        });
    };
    reader.readAsBinaryString(data);
});
userFile2.addEventListener("change", (e) => {
    const data = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (event) => {
        const binaryString = event.target.result;
        workbook = XLSX.read(binaryString, { type: 'binary' });
    };
    reader.readAsBinaryString(data);
});

function generateTable(datas, isImport) {
    tbody.innerHTML = '';
    if (isImport) {
        datas.forEach(data => {
            let newRow = new TabRow(true, data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8])
        });
    } else {

        datas.forEach(data => {
            let complete = true
            let count = 0
            data.forEach(cell => {
                if (cell == 0 || cell == "" || cell == undefined) {
                    complete = false
                } else {
                    count += 1
                }
            })
            if (isNaN(data[3])) {
                complete = false
            }
            if (complete) {
                let newRow = new TabRow(true, "a","a","a","","a", data[0], data[1], data[2], formatNumber(data[3]))
            } else if (count > 0) {
                let newRow = new TabRow(false, undefined, undefined, undefined, undefined, undefined, data[0], data[1], data[2], formatNumber(data[3]))
            }
        });
    }
}

async function loadLotsData() {
    try {
        const response = await fetch('../assets/dpgf.json');
        const data = await response.json();
        return data;
    } catch (error) {
        console.error('Erreur lors du chargement du fichier JSON:', error);
        return [];
    }
}

function formatNumber(number) {
    return Number(number).toLocaleString('fr-FR', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function columnLetterToIndex(letter) {
    let column = 0;
    let length = letter.length;
    for (let i = 0; i < length; i++) {
        column = column * 26 + (letter.charCodeAt(i) - 64);
    }
    return column - 1;
}
readXlsxBtn.addEventListener("click", () => {
    if (!userFile.value || !sheetnameSelect.value || !descriptionInput.value || !uniteInput.value || !quantiteInput.value || !totalInput.value) {
        alert("Veuillez remplir tous les champs obligatoire")
        return
    }
    const worksheet = workbook.Sheets[sheetnameSelect.value]
    const selectedColumns = [descriptionInput.value.toUpperCase(), uniteInput.value.toUpperCase(), quantiteInput.value.toUpperCase(), totalInput.value.toUpperCase()]

    let firstRowNb = firstRow.value

    const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: firstRowNb - 1 });

    const columnIndices = selectedColumns.map(columnLetterToIndex);

    const filteredData = excelData.map(row =>
        columnIndices.map(index => row[index])
    );
    generateTable(filteredData, false);
    initializeLotSelect();
    return
})

readXlsxBtn2.addEventListener("click", () => {
    const worksheet = workbook.Sheets["Tableau de base"]
    const excelData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,  // Utilise des indices numériques au lieu des en-têtes
        defval: "", // Valeur par défaut pour les cellules vides
        raw: false  // Pour obtenir les valeurs formatées plutôt que les valeurs brutes
    });
    initializeLotSelect();
    generateTable(excelData, true)
})

async function initializeLotSelect() {
    const lotsData = await loadLotsData();

    function fillLotSelect(selectElement) {
        selectElement.innerHTML = '<option value="">Sélectionner un lot</option>';
        lotsData.forEach(lotData => {
            const option = document.createElement('option');
            option.value = lotData.lot;
            option.textContent = lotData.lot;
            selectElement.appendChild(option);
        });
    }

    fillLotSelect(lotSelect);

    document.querySelectorAll('.lotSelect').forEach(select => {
        fillLotSelect(select);
    });
}

lotSelect.addEventListener("change", (e) => {
    const selectedLot = e.target.value;

    loadLotsData().then(lotsData => {
        const lotData = lotsData.find(lot => lot.lot === selectedLot);
        sousLotSelect.innerHTML = '<option value="">Sélectionner un sous-lot</option>';

        if (lotData && lotData.sous_elements) {
            lotData.sous_elements.forEach(sousElement => {
                const option = document.createElement('option');
                option.value = sousElement.type;
                option.textContent = sousElement.type;
                sousLotSelect.appendChild(option);
            });
        }
    });

    document.querySelectorAll('.lotSelect').forEach(select => {
        select.value = selectedLot;
        select.dispatchEvent(new Event('change', {
            bubbles: true,
            cancelable: true
        }));
    });
});

sousLotSelect.addEventListener("change", (e) => {
    const selectedSousLot = e.target.value;
    const selectedLot = lotSelect.value;

    document.querySelectorAll('tr').forEach(tr => {
        const lotSelect = tr.querySelector('.lotSelect');
        const sousLotSelect = tr.querySelector('.sousLotSelect');

        if (lotSelect && lotSelect.value === selectedLot) {
            sousLotSelect.value = selectedSousLot;
            sousLotSelect.dispatchEvent(new Event('change', {
                bubbles: true,
                cancelable: true
            }));
        }
    });
});

document.addEventListener("click", (e) => {
    if (e.target.closest("#formTabs")) {
        document.querySelectorAll(".tab").forEach(element => {
            element.classList.toggle("tabSelected")
        });
        document.querySelectorAll(".formContent").forEach(element => {
            element.classList.toggle("hidden")
        })
    }
})

function parseNumberValue(value) {
    if (!value) return 0;
    return parseFloat(value.replace(/\s/g, '').replace(',', '.')) || 0;
}

generateXlsxBtn.addEventListener("click", () => {
    const rows = tbody.querySelectorAll('tr.complete')
    const baseTable = []
    const concaTable = [["Lot", "Sous-lot", "Bloc de base", "Spécificité", "Unité", "Quantité SD", "PU RES", "Montant Total"]]
    rows.forEach(row => {
        let aRow = []
        const cellss = row.querySelectorAll('td')
        cellss.forEach(cell => {
            let a = cell.querySelector("select")
            let b = cell.querySelector("textarea")
            let c = cell.querySelector('input[type="checkbox"')
            if (a) {
                aRow.push(a.value)
            } else if (b) {
                aRow.push(b.value)
            } else if (c) {
                if (c.checked) {
                    aRow.push("true")
                } else {
                    aRow.push("false")
                }
            }
            else {
                aRow.push(cell.textContent)
            }
        });
        baseTable.push(aRow)
        const cells = row.querySelectorAll('td');
        let valid = true
        for (let i = 0; i < cells.length; i++) {
            if ([0, 1, 2].includes(i)) {
                let select = cells[i].querySelector("select")
                if (select.value == "") {
                    valid = false
                }
            }
        }
        if (valid) {
            const rowIdentifier = [
                cells[0].querySelector('select').value,
                cells[1].querySelector('select').value,
                cells[2].querySelector('select').value,
                cells[3].querySelector('textarea').value
            ].join('-');
            const existingRowIndex = concaTable.findIndex(item =>
                item[0] === cells[0].querySelector('select').value &&
                item[1] === cells[1].querySelector('select').value &&
                item[2] === cells[2].querySelector('select').value &&
                item[3] === cells[3].querySelector('textarea').value
            );

            if (existingRowIndex === -1) {
                const newRow = [
                    cells[0].querySelector('select').value,  // lot
                    cells[1].querySelector('select').value,  // sous-lot
                    cells[2].querySelector('select').value,  // bloc de base
                    cells[3].querySelector('textarea').value, // spécificité
                    cells[6].textContent, // Unité
                    cells[4].querySelector('input').checked ? Number(cells[7].textContent) : 0, // quantité
                    Number(cells[7].textContent) / parseNumberValue(cells[8].textContent),
                    parseNumberValue(cells[8].textContent) // total
                ];
                concaTable.push(newRow);
            } else {
                if (cells[4].querySelector('input').checked) {
                    concaTable[existingRowIndex][5] += Number(cells[7].textContent);
                }
                concaTable[existingRowIndex][7] += parseNumberValue(cells[8].textContent);
                concaTable[existingRowIndex][6] = concaTable[existingRowIndex][7] / concaTable[existingRowIndex][5]
            }
        }
    })
    generateXlsx(concaTable, baseTable)
})

function generateXlsx(concaTable, baseTable) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(concaTable);

    const wscols = [
        { wch: 40 }, // Lot
        { wch: 20 }, // Sous-lot
        { wch: 25 }, // Bloc de base
        { wch: 30 }, // Spécificité
        { wch: 8 },  // Unité
        { wch: 12 }, // Quantité
        { wch: 15 }  // Total
    ];
    ws['!cols'] = wscols;

    XLSX.utils.book_append_sheet(wb, ws, "Récapitulatif");

    for (let i = 2; i <= concaTable.length + 1; i++) {
        const totalCell = XLSX.utils.encode_cell({ r: i - 1, c: 6 });
        if (ws[totalCell]) {
            ws[totalCell].z = '#,##0.00 €';
        }
    }
    const ws2 = XLSX.utils.aoa_to_sheet(baseTable);
    XLSX.utils.book_append_sheet(wb, ws2, "Tableau de base")
    XLSX.writeFile(wb, `Tableau_concaténé.xlsx`);
}