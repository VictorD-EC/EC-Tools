const lodTableFile = document.getElementById("lodTableFile");
const versionIFC = document.getElementById("versionIFC");
const sheetName = document.getElementById("sheetName");
const title = document.getElementById("title");
const author = document.getElementById("author");
const version = document.getElementById("version");
const description = document.getElementById("description");
const generateIDSBtn = document.getElementById("generateIDS");

const lodTableFile2 = document.getElementById("lodTableFile2");
const versionIFC2 = document.getElementById("versionIFC2");
const sheetName2 = document.getElementById("sheetName2");
const valuesSheetName = document.getElementById("valuesSheetName")
const title2 = document.getElementById("title2");
const author2 = document.getElementById("author2");
const version2 = document.getElementById("version2");
const description2 = document.getElementById("description2");
const generateIDSBtn2 = document.getElementById("generateIDS2");

function displayRowForm() {
    // Sélectionner les onglets
    const tabs = document.querySelectorAll('#tabs .tab');

    // Mettre à jour les classes des onglets
    tabs[0].classList.add('selected');
    tabs[1].classList.remove('selected');

    // Afficher le formulaire de ligne et cacher celui de colonne
    document.querySelector('#rowForm').classList.remove('hidden');
    document.querySelector('#columnForm').classList.add('hidden');
}

function displayColumnForm() {
    // Sélectionner les onglets
    const tabs = document.querySelectorAll('#tabs .tab');

    // Mettre à jour les classes des onglets
    tabs[0].classList.remove('selected');
    tabs[1].classList.add('selected');

    // Cacher le formulaire de ligne et afficher celui de colonne
    document.querySelector('#rowForm').classList.add('hidden');
    document.querySelector('#columnForm').classList.remove('hidden');
}

// Fonction pour obtenir les données depuis le tableau Excel
function getRowsDatas(worksheet) {
    const data = XLSX.utils.sheet_to_json(worksheet);

    const datas = {};

    // Parcours des lignes
    for (const row of data) {
        const famillesObjet = row["Famille d'objet"];
        const classeIfc = row["Classe IFC"];
        const pset = row["Jeu de propriété Pset"];
        const propriete = row["Propriété"];
        const typeDeValeur = row["Type de valeur"];
        const valeur = row["Valeur"] || null;

        // Si la propriété n'existe pas dans le dictionnaire, on la crée
        if (!datas[propriete]) {
            datas[propriete] = {
                "Pset": pset,
                "TypeDeValeur": typeDeValeur,
                "Entités": [],
                "Valeur": valeur
            };
        }

        // Ajouter la classe IFC à la liste des entités si elle n'y est pas déjà
        if (!datas[propriete]["Entités"].includes(classeIfc)) {
            datas[propriete]["Entités"].push(classeIfc);
        }
    }
    console.log(datas);
    return datas;
}
function getColumnsDatas(worksheet, valuesWorksheet) {
    const data = XLSX.utils.sheet_to_json(worksheet);
    const datas = {};

    const valuesMap = {};
    const valuesData = XLSX.utils.sheet_to_json(valuesWorksheet);
    console.log(valuesData);
    for (const row of valuesData) {
        const propriete = row["Propriété"];
        valuesMap[propriete] = {
            Valeur: row["Valeur"] || null, // Prend la valeur si elle existe, sinon null
            TypeDeValeur: row["Type"] || null, // Prend le type si il existe, sinon null
            Pset: row["Pset"] || null
        };
    }
    console.log(valuesMap);
    for (const row of data) {
        const classeIfc = row["Classe IFC"];
        delete row["Classe IFC"];

        const familleObjet = row["Famille d'objet"];
        delete row["Famille d'objet"];

        for (const propriete in row) {
            if (row[propriete] === "X") {
                if (!datas[propriete]) {
                    datas[propriete] = {
                        Entités: [],
                        Valeur: null,
                        TypeDeValeur: null,
                        Pset: valuesMap[propriete].Pset
                    };
                }
                datas[propriete]["Entités"].push(classeIfc);

                // Ajout de la valeur et du type depuis valuesMap si la propriété existe dans valuesWorksheet
                if (valuesMap[propriete]) {
                    datas[propriete]["Valeur"] = valuesMap[propriete].Valeur;
                    datas[propriete]["TypeDeValeur"] = valuesMap[propriete].TypeDeValeur;
                }
            }
        }
    }
    console.log(datas);
    return datas;
}


// Fonction pour générer le XML en utilisant l'API DOM native
function generateXML(datas, userDatas) {
    // Création du document XML
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString('<xml/>', 'application/xml');

    // Définition des namespaces
    const idsNamespace = "http://standards.buildingsmart.org/IDS";
    const xsNamespace = "http://www.w3.org/2001/XMLSchema";
    const xsiNamespace = "http://www.w3.org/2001/XMLSchema-instance";

    // Fonction utilitaire pour créer des éléments avec le namespace IDS et le préfixe ids:
    function createIdsElement(elementName) {
        return xmlDoc.createElementNS(idsNamespace, elementName); // Pas de préfixe ici
    }

    // Fonction utilitaire pour créer des éléments avec le namespace XML Schema et le préfixe xs:
    function createXsElement(elementName) {
        return xmlDoc.createElementNS(xsNamespace, elementName); // Pas de préfixe ici
    }

    // Créer l'élément racine ids
    const root = xmlDoc.createElementNS(idsNamespace, "ids");
    root.setAttribute("xmlns:ids", idsNamespace);
    root.setAttribute("xmlns:xs", xsNamespace);
    root.setAttribute("xmlns:xsi", xsiNamespace);
    root.setAttributeNS(xsiNamespace, "xsi:schemaLocation", "http://standards.buildingsmart.org/IDS http://standards.buildingsmart.org/IDS/1.0/ids.xsd");

    // Remplacer l'élément racine temporaire <xml/> par l'élément root correct
    xmlDoc.replaceChild(root, xmlDoc.documentElement);

    // Balise info
    const info = createIdsElement("info");
    root.appendChild(info);

    // Création des éléments dans info
    const titleElem = createIdsElement("title");
    titleElem.textContent = userDatas[0];
    info.appendChild(titleElem);

    const authorElem = createIdsElement("author");
    authorElem.textContent = userDatas[1];
    info.appendChild(authorElem);

    if (userDatas[2] !== "") {
        const descriptionElem = createIdsElement("description");
        descriptionElem.textContent = userDatas[2];
        info.appendChild(descriptionElem);
    }

    if (userDatas[3] !== "") {
        const versionElem = createIdsElement("version");
        versionElem.textContent = userDatas[3];
        info.appendChild(versionElem);
    }

    const dateElem = createIdsElement("date");
    dateElem.textContent = new Date().toISOString().split('T')[0];
    info.appendChild(dateElem);

    // Specifications
    const specifications = createIdsElement("specifications");
    root.appendChild(specifications);

    let i = 0;
    for (const [propriete, details] of Object.entries(datas)) {
        i++;
        const specification = createIdsElement("specification");
        specification.setAttribute("name", `Specification ${i}`);
        specification.setAttribute("ifcVersion", userDatas[4]);
        specifications.appendChild(specification);

        // Applicability
        const applicability = createIdsElement("applicability");
        applicability.setAttribute("minOccurs", "1");
        applicability.setAttribute("maxOccurs", "unbounded");
        specification.appendChild(applicability);

        const entity = createIdsElement("entity");
        applicability.appendChild(entity);

        const name = createIdsElement("name");
        entity.appendChild(name);

        const restriction = createXsElement("restriction");
        restriction.setAttribute("base", "xs:string");
        name.appendChild(restriction);

        for (const classeIfc of details["Entités"]) {
            const enumeration = createXsElement("enumeration");
            enumeration.setAttribute("value", String(classeIfc).toUpperCase());
            restriction.appendChild(enumeration);
        }

        // Requirements
        const requirements = createIdsElement("requirements");
        specification.appendChild(requirements);

        // Property
        const propertyElement = createIdsElement("property");
        propertyElement.setAttribute("cardinality", "required");
        requirements.appendChild(propertyElement);

        const propertySet = createIdsElement("propertySet");
        propertyElement.appendChild(propertySet);

        const propertySetSimpleValue = createIdsElement("simpleValue");
        propertySetSimpleValue.textContent = String(details["Pset"]);
        propertySet.appendChild(propertySetSimpleValue);

        const nameElement = createIdsElement("baseName");
        propertyElement.appendChild(nameElement);

        const nameSimpleValue = createIdsElement("simpleValue");
        nameSimpleValue.textContent = propriete;
        nameElement.appendChild(nameSimpleValue);

        // Gestion de la valeur
        const valeur = details["Valeur"];
        if (valeur && String(valeur).toLowerCase() !== "nan") {
            const valeurElement = createIdsElement("value");
            propertyElement.appendChild(valeurElement);

            const xsRestriction = createXsElement("restriction");
            xsRestriction.setAttribute("base", "xs:string");
            valeurElement.appendChild(xsRestriction);

            const valeurs = String(valeur).split(',');
            for (const v of valeurs) {
                const enumeration = createXsElement("enumeration");
                enumeration.setAttribute("value", v.trim());
                xsRestriction.appendChild(enumeration);
            }
        }
    }

    // Conversion en chaîne de caractères
    const serializer = new XMLSerializer();
    const xmlString = serializer.serializeToString(xmlDoc);

    // Formatage du XML (indentation)
    const formattedXml = formatXml(xmlString);

    // Création et téléchargement du fichier
    const fileName = `${userDatas[0].replace(/ /g, '_')}.ids`; // Suppression de la date dans le nom du fichier
    downloadFile(formattedXml, fileName, 'application/xml');

    return formattedXml;
}


// Fonction pour formater le XML (ajouter des indentations)
function formatXml(xml) {
    let formatted = '';
    let indent = '';
    const tab = '  '; // 2 espaces pour l'indentation

    xml.split(/>\s*</).forEach(function (node) {
        if (node.match(/^\/\w/)) { // Balise fermante
            indent = indent.substring(tab.length);
        }

        formatted += indent + '<' + node + '>\r\n';

        if (node.match(/^<?\w[^>]*[^\/]$/) && !node.startsWith("?")) { // Balise ouvrante
            indent += tab;
        }
    });

    return formatted.substring(1, formatted.length - 3);
}

// Fonction pour télécharger le fichier généré
function downloadFile(content, fileName, contentType) {
    const a = document.createElement('a');
    const file = new Blob([content], { type: contentType });
    a.href = URL.createObjectURL(file);
    a.download = fileName;
    a.click();
    URL.revokeObjectURL(a.href);
}



document.addEventListener("click", (e) => {
    const tab = e.target.closest('.tab');
    if (tab) {
        const tabText = tab.querySelector('p').innerText.toLowerCase();

        if (tabText === 'ligne') {
            displayRowForm();
        } else if (tabText === 'colonne') {
            displayColumnForm();
        }
    }
    if (e.target.closest("#generateIDS")) {
        const file = lodTableFile.files[0];
        if (!file) {
            alert("Veuillez sélectionner un fichier Excel contenant le tableau des LOD");
            return;
        }
        if (title.value == "") {
            alert("Veuillez fournir un titre valide pour votre ids");
            return;
        }
        if (author.value == "") {
            alert("Veuillez fournir un auteur valide pour votre ids");
            return;
        }

        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[sheetName.value]

            const titleValue = title.value || "easyIDS";
            const authorValue = author.value || "Eiffage";
            const descriptionValue = description.value || "";
            const versionValue = version.value || "";
            const ifcVersionValue = versionIFC.value || "IFC2x3";

            const userDatas = [titleValue, authorValue, descriptionValue, versionValue, ifcVersionValue];

            const datasFromExcel = getRowsDatas(worksheet);
            generateXML(datasFromExcel, userDatas);
        };

        reader.readAsArrayBuffer(file);
    }
    if (e.target.closest("#generateIDS2")) {
        const file = lodTableFile2.files[0];
        if (!file) {
            alert("Veuillez sélectionner un fichier Excel contenant le tableau des LOD");
            return;
        }
        if (title2.value == "") {
            alert("Veuillez fournir un titre valide pour votre ids");
            return;
        }
        if (author2.value == "") {
            alert("Veuillez fournir un auteur valide pour votre ids");
            return;
        }

        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[sheetName2.value]
            const valuesWorksheet = workbook.Sheets[valuesSheetName.value]

            const titleValue = title2.value || "easyIDS";
            const authorValue = author2.value || "Eiffage";
            const descriptionValue = description2.value || "";
            const versionValue = version2.value || "";
            const ifcVersionValue = versionIFC2.value || "IFC2x3";

            const userDatas = [titleValue, authorValue, descriptionValue, versionValue, ifcVersionValue];

            const datasFromExcel = getColumnsDatas(worksheet, valuesWorksheet);
            generateXML(datasFromExcel, userDatas);
        };

        reader.readAsArrayBuffer(file);
    }
});