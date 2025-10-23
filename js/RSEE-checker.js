const inputFile = document.getElementById("inputFile");
const verifBtn = document.getElementById("verifBtn");
const results = document.getElementById("results")

inputFile.addEventListener("change", (e)=>{
    if (inputFile.files && inputFile.files[0]) {
        verifBtn.className = "ready";
    }
})
verifBtn.addEventListener("click", (e) => {
    if (inputFile.files && inputFile.files[0]) {
        getPaths();
        verifBtn.className = "loading";
    }
})

function getPaths() {
    fetch('../assets/rsee.json')
        .then(response => response.json())
        .then(data => {
            initVerif(data, inputFile.files[0]);
        })
        .catch(error => console.error('Erreur lors du chargement du fichier JSON:', error));
}

function initVerif(paths, file) {
    results.innerHTML = "";

    const reader = new FileReader();

    reader.onload = function (e) {
        const xmlContent = e.target.result;

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlContent, "text/xml");

        let r1 = verifSeul(paths, xmlDoc);
        let r2 = verifBatiment(paths, xmlDoc, "batiment");
        let r3 = verifBatiment(paths, xmlDoc, "Sortie_Batiment_B");
        let r4 = verifBatiment(paths, xmlDoc, "Sortie_Batiment_C");
        let r5 = verifBatiment(paths, xmlDoc, "Sortie_Groupe_D");
        let r6 = verifTypeDonnees(xmlDoc)

        const exist = r1[0] + r2[0] + r3[0] + r4[0] + r5[0];
        const dontExist = r1[1] + r2[1] + r3[1] + r4[1] + r5[1];
        console.log("Il existe " + exist + " balises.");
        console.log("Il manque " + dontExist + " balises.");

        if (dontExist <= 0 && r6 == true) {
            display("✅ Toutes les balises sont présentes et du bon type pour correspondre au PBI !", "valid");
        }
    };

    reader.onerror = function () {
        console.error("Erreur lors de la lecture du fichier");
    };

    reader.readAsText(file);
    inputFile.value = "";
    verifBtn.className = "normal"
}

function verifSeul(paths, xmlDoc) {
    const seuls = paths.seul;
    let exist = 0;
    let dontExist = 0;

    seuls.forEach(balise => {
        const xpathExpression = `//*[name()='${balise.path.split('/').pop()}']`;

        const nodes = xmlDoc.evaluate(xpathExpression, xmlDoc, null, XPathResult.ANY_TYPE, null);
        const node = nodes.iterateNext();

        if (node) {
            console.log(`✅ Le balise ${balise.name} existe. Valeur: ${node.textContent.trim()}`);
            let type = "";
            if (!isNaN(node.textContent.trim())) {
                type = "number";
            } else {
                type = "string";
            }

            if (balise.type == "string") {
                exist += 1;
            } else if (!isNaN(node.textContent.trim()) && balise.type == "number") {
                exist += 1;
            } else {
                display(`⚠️ La balise "${balise.name}" est du type ${type} au lieu de ${balise.type}`, "wrong")
                dontExist += 1;
            }
        } else {
            console.log(`❌ Le balise ${balise.name} n'existe pas.`);
            display(`❌ La balise "${balise.name}" est manquante pour ce projet.`, "error")
            dontExist += 1;
        }
    });

    return [exist, dontExist]
}

function verifBatiment(paths, xmlDoc, elementName) {
    const batiments = xmlDoc.querySelectorAll("projet > Datas_Comp > batiment_collection > batiment");
    const elementPaths = paths[elementName];
    let exist = 0;
    let dontExist = 0;

    for (let index = 1; index < batiments.length + 1; index++) {
        console.log("batiment : " + index);
        elementPaths.forEach(balise => {
            const patternToReplace = `/${elementName}/`;
            const replacement = `/${elementName}[${index + 1}]/`;
            const newPath = balise.path.replace(patternToReplace, replacement);

            const xpathExpression = `//*[name()='${newPath.split('/').pop()}']`;

            const nodes = xmlDoc.evaluate(xpathExpression, xmlDoc, null, XPathResult.ANY_TYPE, null);
            const node = nodes.iterateNext();

            if (node) {
                console.log(`✅ Le balise ${balise.name} existe. Valeur: ${node.textContent.trim()}`);
                let type = "";
                if (!isNaN(node.textContent.trim())) {
                    type = "number";
                } else {
                    type = "string";
                }

                if (balise.type == "string") {
                    exist += 1;
                } else if (!isNaN(node.textContent.trim()) && balise.type == "number") {
                    exist += 1;
                } else {
                    display(`⚠️ La balise "${balise.name}" est du type ${type} au lieu de ${balise.type}.`, "wrong")
                    dontExist += 1;
                }
            } else {
                display(`❌ La balise "${balise.name}" est manquante pour le batiment ${index}.`, "error")
                dontExist += 1;
            }
        });
    }
    return [exist, dontExist]
}

function verifTypeDonnees(xmlDoc) {
    let valid = true;
    const donneesComposant = xmlDoc.querySelectorAll("projet > RSEnv > entree_projet > batiment > zone > contributeur > composant > lot > sous_lot > donnees_composant");
    for (let i = 0; i < donneesComposant.length; i++) {
        const typeDonnees = donneesComposant[i].querySelector('type_donnees');
        if (!typeDonnees) {
            display(`❌ La balise "type_donnees" est manquante dans donnees_composant ${i + 1}.`, "error")
            console.log(`❌ Erreur : balise type_donnees manquante dans donnees_composant ${i + 1}.`);
            valid = false;
        } else {            
            if (typeDonnees.textContent.trim()) {
                
                if (isNaN(typeDonnees.textContent.trim())) {
                    console.log(`La balise "type_donnees" est de type string au lieu de number dans donnees_composant ${i + 1}.`);
                    display(`La balise "type_donnees" est de type string au lieu de number dans donnees_composant ${i + 1}.`, "wrong");
                    valid = false;
                } else {
                    console.log(`donnees_composant ${i + 1} - type_donnees: ${typeDonnees.textContent.trim()}`);
                }
            } else {
                display(`⚠️ La balise "type_donnees" est vide dans donnees_composant ${i + 1}.`, "wrong")
                console.log(`⚠️ Erreur : balise type_donnees est vide dans donnees_composant ${i + 1}.`);
                valid = false
            }
        }

    }

    if (valid) return true;
}

function display(content, type) {
    const div = document.createElement("div");
    switch (type) {
        case "valid":
            div.className = "valid";
            break;
        case "wrong":
            div.className = "wrong";
            break;
        case "error":
            div.className = "error";
            break;
    }
    const p = document.createElement("p");
    p.textContent = content;
    div.appendChild(p);
    results.append(div);
}