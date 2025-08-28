const userFile = document.getElementById('userFile');
const sheetnameSelect = document.getElementById('sheetnameSelect')
const firstRow = document.getElementById('firstRow')
const descriptionInput = document.getElementById('inputDescription')
const uniteInput = document.getElementById('inputUnite')
const quantiteInput = document.getElementById('inputQuantite')
const totalInput = document.getElementById('inputTotal')
const readXlsxBtn = document.getElementById('readXlsx');
const generateXlsxBtn = document.getElementById('generateXlsx');


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
    const worksheet = workbook.Sheets[sheetnameSelect.value]
    const selectedColumns = [descriptionInput.value.toUpperCase(), uniteInput.value.toUpperCase(), quantiteInput.value.toUpperCase(), totalInput.value.toUpperCase()]

    let firstRowNb = firstRow.value

    const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: firstRowNb - 1 });

    const columnIndices = selectedColumns.map(columnLetterToIndex);

    const filteredData = excelData.map(row =>
        columnIndices.map(index => row[index])
    );
    generateTable(filteredData);
    initializeLotSelect();
    return
})

function generateTable(datas) {
    datas.forEach(row => {
        let complete = true
        let count = 0
        let tr = document.createElement('tr')
        row.forEach(cell => {
            if (cell == 0 || cell == "" || cell == undefined) {
                complete = false
            } else {
                count += 1
            }
        })
        if (isNaN(row[3])) {
            complete = false
        }
        if (complete) {
            let lotTd = document.createElement("td")
            let lot = document.createElement("select")
            lot.className = "lotSelect"
            lotTd.appendChild(lot)

            let souslotTd = document.createElement("td")
            let souslot = document.createElement("select")
            souslot.className = "sousLotSelect"
            souslotTd.appendChild(souslot)

            let blocdebaseTd = document.createElement("td")
            let blocdebase = document.createElement("select")
            blocdebase.className = "blocdebase"
            blocdebaseTd.appendChild(blocdebase)

            let specificiteTd = document.createElement("td")
            let specificite = document.createElement("textarea")
            specificiteTd.appendChild(specificite)

            let quantiteBoolTd = document.createElement("td")
            let quantiteBool = document.createElement("input")
            quantiteBool.type = "checkbox"
            quantiteBool.checked = true
            quantiteBoolTd.appendChild(quantiteBool)

            tr.className = "complete"
            tr.append(lotTd, souslotTd, blocdebaseTd, specificiteTd, quantiteBoolTd)
        } else if (count >= 1) {
            let lotTd = document.createElement("td")
            let souslotTd = document.createElement("td")
            let blocdebaseTd = document.createElement("td")
            let specificiteTd = document.createElement("td")
            let quantiteBoolTd = document.createElement("td")

            tr.append(lotTd, souslotTd, blocdebaseTd, specificiteTd, quantiteBoolTd)
        }
        if (complete || count >= 1) {
            for (var i = 0; i < row.length; i++) {
                switch (i) {
                    case 0:
                        let description = document.createElement("td")
                        description.textContent = row[i]
                        tr.appendChild(description)
                        break
                    case 1:
                        let unite = document.createElement("td")
                        unite.textContent = row[i]
                        tr.appendChild(unite)
                        break
                    case 2:
                        let quantite = document.createElement("td")
                        quantite.textContent = row[i]
                        tr.appendChild(quantite)
                        break
                    case 3:
                        let total = document.createElement("td")
                        if (Number(row[i])) {
                            total.textContent = formatNumber(row[i])
                            tr.appendChild(total)
                        } else {
                            tr.appendChild(total)
                        }
                        break
                }
            }
        }
        tbody.appendChild(tr)
    });
}

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

function updateSousLots(selectedLot, targetSelect) {
    loadLotsData().then(lotsData => {
        const lotData = lotsData.find(lot => lot.lot === selectedLot);
        targetSelect.innerHTML = '<option value="">Sélectionner un sous-lot</option>';

        if (lotData && lotData.sous_elements) {
            lotData.sous_elements.forEach(sousElement => {
                const option = document.createElement('option');
                option.value = sousElement.type;
                option.textContent = sousElement.type;
                targetSelect.appendChild(option);
            });
        }
    });
}

function updateBlocsDeBase(selectedLot, selectedSousLot, targetSelect) {
    loadLotsData().then(lotsData => {
        const lotData = lotsData.find(lot => lot.lot === selectedLot);
        targetSelect.innerHTML = '<option value="">Sélectionner un bloc de base</option>';

        if (lotData && lotData.sous_elements) {
            const sousElement = lotData.sous_elements.find(se => se.type === selectedSousLot);
            if (sousElement && sousElement.elements) {
                sousElement.elements.forEach(element => {
                    const option = document.createElement('option');
                    option.value = element;
                    option.textContent = element;
                    targetSelect.appendChild(option);
                });
            }
        }
    });
}

lotSelect.addEventListener("change", (e) => {
    const selectedLot = e.target.value;

    document.querySelectorAll('.lotSelect').forEach(select => {
        select.value = selectedLot;
    });

    updateSousLots(selectedLot, sousLotSelect);

    document.querySelectorAll('.sousLotSelect').forEach(select => {
        updateSousLots(selectedLot, select);
    });
});

sousLotSelect.addEventListener("change", (e) => {
    const selectedSousLot = e.target.value;
    const selectedLot = lotSelect.value;

    document.querySelectorAll('tr').forEach(tr => {
        const lotSelect = tr.querySelector('.lotSelect');
        const sousLotSelect = tr.querySelector('.sousLotSelect');
        const blocDeBase = tr.querySelector('.blocdebase');

        if (lotSelect && lotSelect.value === selectedLot) {
            sousLotSelect.value = selectedSousLot;
            updateBlocsDeBase(selectedLot, selectedSousLot, blocDeBase);
        }
    });
});

document.addEventListener('change', (e) => {
    if (e.target.classList.contains('lotSelect')) {
        const tr = e.target.closest('tr');
        const sousLotSelect = tr.querySelector('.sousLotSelect');
        const blocDeBase = tr.querySelector('.blocdebase');
        updateSousLots(e.target.value, sousLotSelect);
        blocDeBase.innerHTML = '<option value="">Sélectionner un bloc de base</option>';
    }

    if (e.target.classList.contains('sousLotSelect')) {
        const tr = e.target.closest('tr');
        const lotSelect = tr.querySelector('.lotSelect');
        const blocDeBase = tr.querySelector('.blocdebase');
        updateBlocsDeBase(lotSelect.value, e.target.value, blocDeBase);
    }
});

function parseNumberValue(value) {
    if (!value) return 0;
    return parseFloat(value.replace(/\s/g, '').replace(',', '.')) || 0;
}

generateXlsxBtn.addEventListener("click", () => {
    const rows = tbody.querySelectorAll('tr.complete')
    const concaTable = [["Lot", "Sous-lot", "Bloc de base", "Spécificité", "Unité", "Quantité", "Montant Total"]]
    rows.forEach(row => {
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
            // Créer l'identifiant unique basé sur les colonnes 0,1,2,3
            const rowIdentifier = [
                cells[0].querySelector('select').value,
                cells[1].querySelector('select').value,
                cells[2].querySelector('select').value,
                cells[3].querySelector('textarea').value
            ].join('-');

            // Chercher si la ligne existe déjà dans concaTable
            const existingRowIndex = concaTable.findIndex(item =>
                item[0] === cells[0].querySelector('select').value &&
                item[1] === cells[1].querySelector('select').value &&
                item[2] === cells[2].querySelector('select').value &&
                item[3] === cells[3].querySelector('textarea').value
            );

            if (existingRowIndex === -1) {
                // Si la ligne n'existe pas, créer une nouvelle entrée
                const newRow = [
                    cells[0].querySelector('select').value,  // lot
                    cells[1].querySelector('select').value,  // sous-lot
                    cells[2].querySelector('select').value,  // bloc de base
                    cells[3].querySelector('textarea').value, // spécificité
                    cells[6].textContent, // Unité
                    cells[4].querySelector('input').checked ? Number(cells[7].textContent) : 0, // quantité
                    parseNumberValue(cells[8].textContent) // total
                ];
                concaTable.push(newRow);
            } else {
                // Si la ligne existe déjà
                if (cells[4].querySelector('input').checked) {
                    // Ajouter la quantité si checkbox cochée
                    concaTable[existingRowIndex][5] += Number(cells[7].textContent);
                }
                // Ajouter le total dans tous les cas
                concaTable[existingRowIndex][6] += parseNumberValue(cells[8].textContent);
            }
        }
    })
    generateXlsx(concaTable)
})

function generateXlsx(concaTable) {
    // Créer un nouveau classeur
    const wb = XLSX.utils.book_new();
    // Créer une feuille à partir des données
    const ws = XLSX.utils.aoa_to_sheet(concaTable);

    // Définir des styles de cellule (largeur des colonnes)
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

    // Ajouter la feuille au classeur
    XLSX.utils.book_append_sheet(wb, ws, "Récapitulatif");

    // Formater les nombres dans les colonnes de quantité et total
    // Commencer à partir de la ligne 2 (après les en-têtes)
    for (let i = 2; i <= concaTable.length + 1; i++) {

        // Formatage du total (colonne G)
        const totalCell = XLSX.utils.encode_cell({ r: i - 1, c: 6 });
        if (ws[totalCell]) {
            ws[totalCell].z = '#,##0.00 €';
        }
    }

    // Générer le fichier et lancer le téléchargement
    XLSX.writeFile(wb, `Tableau_concaténé.xlsx`);
}