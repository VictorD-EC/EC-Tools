document.addEventListener('DOMContentLoaded', () => {
    const pdfFileInput = document.getElementById('pdfFile');
    const resultDisplay = document.getElementById('resultDisplay');

    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

    // ---- Définition des valeurs attendues pour l'image ----
    const expected = {
        id: 'img_p0_1',
        x: 47.45,
        y: 768.00,
        width: 122.90,
        height: 40.25,
    };
    const tolerance = 0.50; // Tolérance pour la comparaison des nombres flottants

    // Fonction pour afficher les messages sur la page
    function displayMessage(message, type = 'info') {
        resultDisplay.innerHTML = `<p>${message}</p>`;
        resultDisplay.className = `resultDisplay ${type}`;
    }

    pdfFileInput.addEventListener('change', async (event) => {
        resultDisplay.innerHTML = '';
        resultDisplay.className = 'resultDisplay';

        const file = event.target.files[0];
        if (!file) {
            console.log('Aucun fichier sélectionné.');
            displayMessage("Veuillez sélectionner un fichier PDF.", "warning");
            return;
        }

        displayMessage("Analyse du PDF en cours...", "info");
        const fileReader = new FileReader();
        fileReader.onload = async () => {
            const typedarray = new Uint8Array(fileReader.result);

            try {
                const loadingTask = pdfjsLib.getDocument(typedarray);
                const pdf = await loadingTask.promise;

                console.log(`✅ PDF chargé : ${pdf.numPages} page(s).`);

                const page = await pdf.getPage(1);
                const viewport = page.getViewport({ scale: 1 });

                const opList = await page.getOperatorList();

                let foundFirstImage = false;

                for (let i = 0; i < opList.fnArray.length; i++) {
                    const fn = opList.fnArray[i];
                    const args = opList.argsArray[i];

                    if (fn === pdfjsLib.OPS.save) {
                        let transformMatrix = null;
                        let imageName = null;

                        for (let j = i + 1; j < opList.fnArray.length; j++) {
                            const nextFn = opList.fnArray[j];
                            const nextArgs = opList.argsArray[j];

                            if (nextFn === pdfjsLib.OPS.restore) {
                                break;
                            }

                            if (nextFn === pdfjsLib.OPS.transform) {
                                transformMatrix = nextArgs;
                            }

                            if (nextFn === pdfjsLib.OPS.paintImageXObject ||
                                nextFn === pdfjsLib.OPS.paintJpegXObject ||
                                nextFn === pdfjsLib.OPS.paintImageMask) {
                                imageName = nextArgs[0];
                                break;
                            }
                        }

                        if (transformMatrix && imageName) {
                            const [a, b, c, d, e, f] = transformMatrix;

                            const x = e;
                            const y = f;
                            const width = a;
                            const height = d;

                            console.log(`\n💡 Première image trouvée (ID: ${imageName}) :`);
                            console.log(`   - Coordonnées (unité PDF, coin inf-gauche) : X=${x.toFixed(2)}, Y=${y.toFixed(2)}`);
                            console.log(`   - Taille (unité PDF) : Largeur=${width.toFixed(2)}, Hauteur=${height.toFixed(2)}`);
                            console.log(`   - Matrice de transformation : [${transformMatrix.map(n => n.toFixed(2)).join(', ')}]`);

                            foundFirstImage = true;

                            // ---- Logique de vérification simplifiée ----
                            let isMatch = true;
                            let details = [];
                            let locationMismatch = false;
                            let sizeMismatch = false;

                            if (imageName !== expected.id) {
                                isMatch = false;
                                details.push(`ID de l'image incorrect (Attendu: ${expected.id}, Obtenu: ${imageName})`);
                            }

                            if (Math.abs(x - expected.x) > tolerance || Math.abs(y - expected.y) > tolerance) {
                                isMatch = false;
                                locationMismatch = true;
                            }

                            if (Math.abs(width - expected.width) > tolerance || Math.abs(height - expected.height) > tolerance) {
                                isMatch = false;
                                sizeMismatch = true;
                            }

                            if (locationMismatch) {
                                details.push(`Mauvais emplacement (Attendu: X=${expected.x.toFixed(2)}, Y=${expected.y.toFixed(2)} - Obtenu: X=${x.toFixed(2)}, Y=${y.toFixed(2)})`);
                            }
                            if (sizeMismatch) {
                                details.push(`Mauvaise taille (Attendue: L=${expected.width.toFixed(2)}, H=${expected.height.toFixed(2)} - Obtenue: L=${width.toFixed(2)}, H=${height.toFixed(2)})`);
                            }


                            if (isMatch) {
                                displayMessage("✅ La première image correspond parfaitement aux critères attendus !", "valid");
                            } else {
                                displayMessage(`❌ La première image NE CORRESPOND PAS entièrement aux critères attendus.<br>${details.join('<br>')}`, "error");
                            }

                            break; // Arrêter après la première image trouvée et vérifiée
                        }
                    }
                }

                if (!foundFirstImage) {
                    displayMessage("⚠️ Aucune image n'a été détectée sur la première page de ce PDF avec l'approche actuelle.", "warning");
                    console.warn("   L'analyse des opérateurs de PDF peut être complexe et dépendre de la façon dont le PDF a été généré.");
                }

            } catch (error) {
                console.error("❌ Erreur lors du chargement ou de l'analyse du PDF :", error);
                displayMessage(`❌ Erreur lors de l'analyse du PDF : ${error.message}. Veuillez vérifier la console pour plus de détails.`, "failure");
            }
        };
        fileReader.readAsArrayBuffer(file);
    });
});