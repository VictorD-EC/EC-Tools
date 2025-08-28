document.addEventListener('DOMContentLoaded', async () => {
    const userFile = document.getElementById('userFile');
    const readXlsxBtn = document.getElementById('readXlsx');
    const generateXlsxBtn = document.getElementById('generateXlsx'); // Ajout de la référence au bouton "Générer Excel"
    const mergedTableContainer = document.getElementById('mergedTable');
    const lotSelect = document.getElementById('lotSelect');
    const sousLotSelect = document.getElementById('sousLotSelect');
    let workbook = []

    document.addEventListener("click", (e) => {
        if (e.target.matches('.concatenable')) {
            e.target.classList.toggle("conca")
        }
    })

    // Fonction pour convertir une valeur en nombre si possible
    function convertToNumber(value) {
        if (typeof value === 'number') {
            return value; // Déjà un nombre
        }
        if (typeof value === 'string') {
            // Supprimer les séparateurs de milliers et remplacer la virgule par un point
            const cleanedValue = value.replace(/[\s ]/g, '').replace(',', '.');
            const num = Number(cleanedValue);
            if (!isNaN(num)) {
                return num; // Conversion réussie
            }
        }
        return NaN; // Retourne NaN si la conversion échoue
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

    function updateBlocBaseOptions(select, lotsData, selectedLot, selectedSousLot) {
        select.innerHTML = '<option value="">Sélectionner un bloc de base</option>'; // Réinitialiser les options

        const selectedLotData = lotsData.find(lot => lot.lot === selectedLot);
        if (selectedLotData) {
            const selectedSousLotData = selectedLotData.sous_elements.find(sousElement => sousElement.type === selectedSousLot);
            if (selectedSousLotData && selectedSousLotData.elements) {
                selectedSousLotData.elements.forEach(element => {
                    const option = document.createElement('option');
                    option.value = element;
                    option.textContent = element;
                    select.appendChild(option);
                });
                select.disabled = false;
            } else {
                select.disabled = true;
            }
        } else {
            select.disabled = true;
        }
    }

    // Initialisation des options du sélecteur de lot global
    async function initializeLotSelect(lotsData) {
        lotsData.forEach(lot => {
            const option = document.createElement('option');
            option.value = lot.lot;
            option.textContent = lot.lot;
            lotSelect.appendChild(option);
        });

        // Gestion du changement de lot global
        lotSelect.addEventListener('change', () => {
            const selectedLot = lotSelect.value;

            // Mettre à jour les sélecteurs de lot dans chaque ligne
            const lotSelects = mergedTableContainer.querySelectorAll('tbody td:nth-child(1) select');
            lotSelects.forEach(select => {
                select.value = selectedLot;
                updateSousLotOptions(select.parentNode.nextElementSibling.querySelector('select'), lotsData, selectedLot);
            });

            // Mettre à jour le sélecteur de sous-lot global
            updateSousLotSelect(sousLotSelect, lotsData, selectedLot);

            // Mettre à jour les sélecteurs de sous-lot dans chaque ligne
            const sousLotSelects = mergedTableContainer.querySelectorAll('tbody td:nth-child(2) select');
            sousLotSelects.forEach(select => {
                updateSousLotOptions(select, lotsData, selectedLot);
            });
        });
        sousLotSelect.addEventListener('change', () => {
            const selectedSousLot = sousLotSelect.value;

            // Cibler les sélecteurs de sous-lot (2ème colonne) et non les blocs de base (3ème colonne)
            const sousLotSelects = mergedTableContainer.querySelectorAll('tbody td:nth-child(2) select');

            sousLotSelects.forEach(select => {
                // Vérifier si la valeur sélectionnée existe dans les options du sélecteur
                const optionExists = Array.from(select.options).some(option => option.value === selectedSousLot);

                // Si l'option existe, définir la valeur du sélecteur
                if (optionExists) {
                    select.value = selectedSousLot;

                    // Mettre à jour le bloc de base correspondant à cette ligne
                    const blocBaseSelect = select.parentNode.nextElementSibling.querySelector('select');
                    if (blocBaseSelect) {
                        updateBlocBaseOptions(blocBaseSelect, lotsData, select.parentNode.previousElementSibling.querySelector('select').value, selectedSousLot);
                    }
                }
            });
        });
    }

    // Fonction pour mettre à jour les options du sélecteur de sous-lot global
    function updateSousLotSelect(select, lotsData, selectedLot) {
        select.innerHTML = '<option value="">Sélectionner un sous-lot</option>'; // Réinitialiser les options

        const selectedLotData = lotsData.find(lot => lot.lot === selectedLot);
        if (selectedLotData && selectedLotData.sous_elements) {
            selectedLotData.sous_elements.forEach(sousElement => {
                const option = document.createElement('option');
                option.value = sousElement.type;
                option.textContent = sousElement.type;
                select.appendChild(option);
            });
            select.disabled = false; // Activer le sélecteur de sous-lot
        } else {
            select.disabled = true; // Désactiver le sélecteur de sous-lot
        }
    }

    async function displayMergedTable(data, lotsData) {
        mergedTableContainer.innerHTML = '';

        // Créer les en-têtes combinés
        const staticHeaders = ["Lot", "Sous-lot", "Bloc de base", "Spécificités", "Quantité?"];
        const excelHeaders = Object.keys(data[0]).filter(header =>
            !staticHeaders.includes(header)); // Obtenir uniquement les en-têtes Excel uniques
        const allHeaders = [...staticHeaders, ...excelHeaders];

        // Créer l'en-tête du tableau
        let thead = document.createElement('thead');
        let headerRow = document.createElement('tr');
        console.log(staticHeaders);
        allHeaders.forEach(headerText => {
            let header = document.createElement('th');
            header.textContent = headerText || '';
            header.classList.add("conca")
            console.log(staticHeaders.indexOf(headerText));
            if (staticHeaders.indexOf(headerText) == "-1") {
                header.classList.add("concatenable")
            }
            headerRow.appendChild(header);
        });

        thead.appendChild(headerRow);
        mergedTableContainer.appendChild(thead);

        // Créer le corps du tableau
        let tbody = document.createElement('tbody');

        data.forEach(rowData => {
            let row = document.createElement('tr');

            // Cellule Lot
            let lotCell = document.createElement('td');
            lotCell.classList.add("selectCell")
            let lotSelect = document.createElement('select');
            lotsData.forEach(lot => {
                const option = document.createElement('option');
                option.value = lot.lot;
                option.textContent = lot.lot;
                lotSelect.appendChild(option);
            });
            lotSelect.value = rowData.Lot || lotSelect.firstChild?.value || '';
            lotSelect.addEventListener('change', () => {
                updateSousLotOptions(sousLotSelect, lotsData, lotSelect.value);
            });
            lotCell.appendChild(lotSelect);
            row.appendChild(lotCell);

            // Cellule Sous-lot
            let sousLotCell = document.createElement('td');
            sousLotCell.classList.add("selectCell")
            let sousLotSelect = document.createElement('select');
            updateSousLotOptions(sousLotSelect, lotsData, lotSelect.value);
            sousLotSelect.value = rowData["Sous-lot"] || sousLotSelect.firstChild?.value || '';
            sousLotCell.appendChild(sousLotSelect);
            row.appendChild(sousLotCell);

            // Cellule Bloc de base
            let blocBaseCell = document.createElement('td');
            blocBaseCell.classList.add("selectCell")
            let blocBaseSelect = document.createElement('select');
            const selectedLotData = lotsData.find(lot => lot.lot === lotSelect.value);
            let blocBaseOptions = [];
            if (selectedLotData) {
                selectedLotData.sous_elements.forEach(sousElement => {
                    if (sousElement.type === sousLotSelect.value) {
                        blocBaseOptions = sousElement.elements;
                    }
                });
            }
            blocBaseOptions.forEach(blocBase => {
                const option = document.createElement('option');
                option.value = blocBase;
                option.textContent = blocBase;
                blocBaseSelect.appendChild(option);
            });
            blocBaseCell.appendChild(blocBaseSelect);
            row.appendChild(blocBaseCell);

            // Cellules Spécificités
            ["Spécificités"].forEach(fieldName => {
                let cell = document.createElement('td');
                cell.classList.add("textareaCell")
                let textarea = document.createElement('textarea');
                textarea.value = rowData[fieldName] || '';
                cell.appendChild(textarea);
                row.appendChild(cell);
            });

            // Cellules Quantité
            ["Quantité?"].forEach(fieldName => {
                let cell = document.createElement('td');
                cell.classList.add("inputCheckboxCell")
                let input = document.createElement('input');
                input.type = 'checkbox';
                input.checked = 'true'
                input.value = rowData[fieldName] || '';
                cell.appendChild(input);
                row.appendChild(cell);
            });

            // Ajouter les cellules pour les données Excel
            excelHeaders.forEach(header => {
                let cell = document.createElement('td');
                let value = rowData[header] || '';

                // Vérifier si la valeur est un objet et extraire la propriété 'result'
                if (typeof value === 'object' && value !== null && value.hasOwnProperty('result')) {
                    value = value.result;
                }

                if (typeof value === 'number') { // Vérifie si la valeur est un nombre
                    value = value.toLocaleString('fr-FR', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2,
                        useGrouping: true
                    });
                }

                cell.textContent = value;
                row.appendChild(cell);
            });

            tbody.appendChild(row);
        });

        mergedTableContainer.appendChild(tbody);
    }



    // Fonction pour mettre à jour les options du sélecteur de sous-lot dans une ligne
    function updateSousLotOptions(select, lotsData, selectedLot) {
        select.innerHTML = '<option value="">Sélectionner un sous-lot</option>'; // Réinitialiser les options

        const selectedLotData = lotsData.find(lot => lot.lot === selectedLot);
        if (selectedLotData && selectedLotData.sous_elements) {
            selectedLotData.sous_elements.forEach(sousElement => {
                const option = document.createElement('option');
                option.value = sousElement.type;
                option.textContent = sousElement.type;
                select.appendChild(option);
            });
            select.disabled = false; // Activer le sélecteur de sous-lot
        } else {
            select.disabled = true; // Désactiver le sélecteur de sous-lot
        }
    }

    async function mergeData(staticData, excelData, lotsData) {
        const excelHeaders = excelData[0];
        excelData.shift();

        const excelDataObjects = excelData.map(row => {
            const obj = {};
            excelHeaders.forEach((header, index) => {
                obj[header] = row[index] || '';
            });
            return obj;
        });

        // Créer un tableau de la même longueur que les données Excel
        staticData = new Array(excelDataObjects.length).fill({
            "Lot": "",
            "Sous-lot": "",
            "Bloc de base": "",
            "Spécificités": "",
            "Quantité?": ""
        });

        return excelDataObjects.map((excelRow, index) => ({
            ...staticData[index],
            ...excelRow
        }));
    }

    readXlsxBtn.addEventListener('click', async () => {
        const file = userFile.files[0];
        if (!file) {
            alert('Veuillez sélectionner un fichier Excel.');
            return;
        }

        const sheetNameInput = document.getElementById('sheetName');
        if (!sheetNameInput.value) {
            alert('Veuillez entrer le nom de la feuille.');
            return;
        }
        const sheetName = sheetNameInput.value;

        const topLeftColumnInput = document.getElementById('topLeftColumn');
        const topLeftRowInput = document.getElementById('topLeftRow');
        const bottomRightColumnInput = document.getElementById('bottomRightColumn');
        const bottomRightRowInput = document.getElementById('bottomRightRow');

        const topLeftColumn = topLeftColumnInput.value.toUpperCase();
        const topLeftRow = parseInt(topLeftRowInput.value);
        const bottomRightColumn = bottomRightColumnInput.value.toUpperCase();
        const bottomRightRow = parseInt(bottomRightRowInput.value);

        if (!topLeftColumn || !topLeftRow || !bottomRightColumn || !bottomRightRow) {
            alert('Veuillez entrer les coordonnées des cases haut gauche et bas droite.');
            return;
        }

        try {
            const arrayBuffer = await file.arrayBuffer();
            workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);

            const worksheet = workbook.getWorksheet(sheetName);
            if (!worksheet) {
                alert(`La feuille "${sheetName}" n'existe pas dans le fichier.`);
                return;
            }

            // Convertir les colonnes en indices
            const startCol = columnNameToNumber(topLeftColumn);
            const endCol = columnNameToNumber(bottomRightColumn);

            // Extraire les données
            const excelData = [];
            const headers = [];

            // Extraction des en-têtes
            worksheet.getRow(topLeftRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
                if (colNumber >= startCol && colNumber <= endCol) {
                    headers.push(cell.value);
                }
            });
            excelData.push(headers);

            // Extraction des données
            for (let row = topLeftRow + 1; row <= bottomRightRow; row++) {
                const rowData = [];
                worksheet.getRow(row).eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    if (colNumber >= startCol && colNumber <= endCol) {
                        rowData.push(cell.value);
                    }
                });
                if (rowData.length > 0) excelData.push(rowData);
            }

            const lotsData = await loadLotsData();
            const mergedData = await mergeData([], excelData, lotsData);
            await displayMergedTable(mergedData, lotsData);
            workbook.mergedData = mergedData;
            workbook.lotsData = lotsData;
            workbook.sheetName = sheetName;

        } catch (error) {
            console.error('Erreur lors de la lecture du fichier Excel:', error);
            alert('Erreur lors de la lecture du fichier Excel');
        }
    });

    // Fonction utilitaire pour convertir le nom de colonne en nombre
    function columnNameToNumber(columnName) {
        let result = 0;
        for (let i = 0; i < columnName.length; i++) {
            result *= 26;
            result += columnName.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
        }
        return result;
    }

    // Charger les données et initialiser les sélecteurs
    const lotsData = await loadLotsData();
    await initializeLotSelect(lotsData);


    generateXlsxBtn.addEventListener('click', async () => {
        const datas = [];
        const finalData = []; // Tableau final concaténé

        const headers = Array.from(document.querySelectorAll('#mergedTable thead th')).map(th => {
            return {
                text: th.textContent,
                conca: th.classList.contains('conca')
            };
        });

        // Récupérer les lignes complètes du tableau
        const rows = Array.from(document.querySelectorAll('#mergedTable tbody tr'));
        rows.forEach((row, rowIndex) => {
            let isRowComplete = true;
            let arow = [];

            const cells = Array.from(row.querySelectorAll('td'));
            cells.forEach((cell, cellIndex) => {
                switch (cell.className) {
                    case "selectCell":
                        const select = cell.querySelector("select");
                        if (!select) {
                            isRowComplete = false;
                            return; // Sortir de cells.forEach si la ligne est incomplete
                        }
                        const selectCell = select.value;
                        if (selectCell === "") {
                            isRowComplete = false;
                        } else {
                            arow.push(selectCell);
                        }
                        break;

                    case "textareaCell":
                        const textareaCell = cell.querySelector("textarea");
                        if (!textareaCell) {
                            isRowComplete = false;
                            console.log('text');
                        } else {
                            arow.push(textareaCell.value);
                        }
                        break;

                    case "inputCheckboxCell":
                        const inputCheckboxCell = cell.querySelector('input[type="checkbox"]');
                        if (!inputCheckboxCell) {
                            isRowComplete = false;
                        } else {
                            arow.push(inputCheckboxCell.checked);
                        }
                        break;

                    default:
                        const cellText = cell.textContent;
                        if (cellText === "") {
                            isRowComplete = false;
                        } else {
                            arow.push(cellText);
                        }
                        break;
                }

                if (!isRowComplete) {
                    return; // Sortir de cells.forEach si la ligne est incomplete
                }
            });

            if (isRowComplete) {
                datas.push(arow);
            }
        });

        // console.log("Headers:", headers);
        // console.log("Raw Data:", datas);

        // Concaténation des données
        datas.forEach(dataRow => {
            const lot = dataRow[0];         // Assumes Lot is always at index 0
            const sousLot = dataRow[1];      // Assumes Sous-Lot is always at index 1
            const blocBase = dataRow[2];     // Assumes Bloc de base is always at index 2
            const specificites = dataRow[3];  // Assumes Spécificités is always at index 3

            // Crée une clé unique pour identifier la ligne
            const existingRowIndex = finalData.findIndex(row => row.lot === lot && row.sousLot === sousLot && row.blocBase === blocBase && row.specificites === specificites);

            if (existingRowIndex === -1) {
                // La ligne n'existe pas, on la crée et l'ajoute
                const newRow = {
                    lot: lot,
                    sousLot: sousLot,
                    blocBase: blocBase,
                    specificites: specificites,
                };

                for (let i = 4; i < headers.length; i++) {
                    if (headers[i].conca) {  // Ignore les colonnes avec conca: false
                        const headerText = headers[i].text.replace(/\s/g, ''); // Remove spaces from header text
                        const cellValue = dataRow[i];

                        // Initialiser la valeur avec la valeur de la case
                        const cellNumber = convertToNumber(cellValue);
                        if (!isNaN(cellNumber)) {
                            newRow[headerText] = cellNumber; //Initialiser à un nombre
                        } else {
                            newRow[headerText] = cellValue; // Sinon, copier la valeur directement
                        }
                    }
                }


                finalData.push(newRow);
            } else {
                // La ligne existe déjà, on met à jour les valeurs à partir de la colonne 4
                console.log("Ligne existante détectée, mise à jour des valeurs...");
                for (let i = 4; i < headers.length; i++) {
                    if (headers[i].conca) {  // Ignore les colonnes avec conca: false
                        const headerText = headers[i].text.replace(/\s/g, ''); // Remove spaces from header text
                        console.log(`Colonne ${i}: ${headerText}`);

                        // Trouver l'index correct de la colonne "Quantité"
                        const quantiteIndex = headers.findIndex(header => header.text === "Quantité");
                        console.log(`Index de la colonne "Quantité": ${quantiteIndex}`);

                        if (i === quantiteIndex) {
                            console.log("C'est la colonne quantité");
                            // Vérifier si la checkbox est cochée (dataRow[4] est la valeur de la checkbox)
                            if (dataRow[4]) {
                                console.log("Checkbox cochée");
                                // Si c'est coché, on additionne
                                const cellValue = dataRow[i];
                                console.log(`Valeur de la cellule Quantité: ${cellValue}`);
                                const cellNumber = convertToNumber(cellValue);
                                console.log(`Valeur convertie en nombre: ${cellNumber}`);

                                if (!isNaN(cellNumber)) {
                                    console.log("C'est un nombre valide, on additionne");
                                    finalData[existingRowIndex][headerText] = Number(finalData[existingRowIndex][headerText] || 0) + cellNumber;
                                    console.log(`Nouvelle valeur de finalData[${existingRowIndex}][${headerText}]: ${finalData[existingRowIndex][headerText]}`);
                                } else {
                                    console.log("Ce n'est pas un nombre valide.");
                                }
                            } else {
                                console.log("Checkbox pas cochée.");
                            }
                        } else {
                            console.log("Pas la colonne quantité, valeur : " + dataRow[i]);
                            // Si ce n'est pas la colonne "Quantité", on vérifie si c'est un nombre
                            const cellValue = dataRow[i];
                            const cellNumber = convertToNumber(cellValue);
                            console.log(`Valeur convertie en nombre: ${cellNumber}`);
                            if (!isNaN(cellNumber)) {
                                console.log("C'est un nombre");
                                // Si c'est un nombre, on l'additionne
                                const currentValue = Number(finalData[existingRowIndex][headerText] || 0);
                                console.log(finalData[existingRowIndex]);
                                console.log(headerText);
                                console.log(finalData[existingRowIndex][headerText]);
                                console.log(currentValue);
                                finalData[existingRowIndex][headerText] = currentValue + cellNumber;
                                console.log(`Nouvelle valeur de finalData[${existingRowIndex}][${headerText}]: ${finalData[existingRowIndex][headerText]}`);
                            } else {
                                console.log("Ce n'est pas un nombre, on ne fait rien.");
                            }
                        }
                    }
                }
            }

        });

        console.log("Final Data:", finalData);

        // Création de la feuille excel avec le tableau final concaténé
        if (finalData.length > 0) {
            try {
                // Créer un nouveau workbook
                const wb = workbook
                const ws = wb.addWorksheet('ConcatenedData');

                // Ajouter les en-têtes
                const headerRow = ws.addRow(Object.keys(finalData[0]));
                headerRow.font = { bold: true };

                // Ajouter les données
                finalData.forEach(data => {
                    const row = ws.addRow(Object.values(data));
                });

                // Ajuster la largeur des colonnes
                ws.columns.forEach(column => {
                    column.width = 15;
                });

                // Générer le fichier
                const buffer = await wb.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'concatened_data.xlsx';
                a.click();
                window.URL.revokeObjectURL(url);

            } catch (error) {
                console.error('Erreur lors de la génération du fichier Excel:', error);
                alert('Erreur lors de la génération du fichier Excel');
            }
        } else {
            console.log("Aucune donnée à exporter.");
        }
    });
});