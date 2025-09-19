export class TabRow {
    constructor(isComplete, lot, sousLot, blocDeBase, specificite, quantiteBool, description, unite, quantite, montantTotal) {
        this.isComplete = isComplete;
        this.lot = lot;
        this.sousLot = sousLot;
        this.blocDeBase = blocDeBase;
        this.specificite = specificite;
        this.quantiteBool = quantiteBool;
        this.description = description;
        this.unite = unite;
        this.quantite = quantite;
        this.montantTotal = montantTotal;

        this.tbody = document.querySelector("tbody")

        this.loadData().then(() => {
            this.initRow();
        });
    }

    async loadData() {
        try {
            const response = await fetch('../assets/dpgf.json');
            if (!response.ok) {
                throw new Error(`Erreur HTTP! Statut: ${response.status}`);
            }
            this.lotsData = await response.json();
        } catch (error) {
            console.error("Erreur lors de la récupération du JSON:", error);
        }
    }

    initRow() {
        const tr = document.createElement("tr");

        if (this.isComplete) {
            tr.className = 'complete'
        }

        const lotTd = document.createElement("td");
        const sousLotTd = document.createElement("td");
        const blocDeBaseTd = document.createElement("td");

        if (this.lot && this.sousLot && this.blocDeBase) {
            const lotSelect = document.createElement("select");
            lotSelect.className = "lotSelect";
            lotSelect.innerHTML = '<option value="">Sélectionner un lot</option>'
            this.lotsData.forEach(lotData => {
                const option = document.createElement("option");
                option.value = lotData.lot;
                option.textContent = lotData.lot;
                lotSelect.appendChild(option)
            });
            setTimeout(() => {
                if (this.lot && this.lotsData.some(lotData => lotData.lot === this.lot)) {
                    lotSelect.value = this.lot;
                    lotSelect.dispatchEvent(new Event('change', {
                        bubbles: true,
                        cancelable: true
                    }));
                }
            }, 0);
            lotTd.appendChild(lotSelect)

            lotSelect.addEventListener("change", () => {
                sousLotSelect.innerHTML = '';
                blocDeBaseSelect.innerHTML = '';
                sousLotSelect.innerHTML = '<option value="">Sélectionner un sous-lot</option>';
                this.lot = lotSelect.value;
                const lotData = this.lotsData.find(lot => lot.lot === this.lot);
                lotData.sous_elements.forEach(sousElement => {
                    const option = document.createElement("option");
                    option.value = sousElement.type;
                    option.textContent = sousElement.type;
                    sousLotSelect.appendChild(option);
                });
                if (this.sousLot && this.lot) {
                    const lotData = this.lotsData.find(lot => lot.lot === this.lot);
                    if (lotData && lotData.sous_elements.some(sousElement => sousElement.type === this.sousLot)) {
                        sousLotSelect.value = this.sousLot;
                        sousLotSelect.dispatchEvent(new Event('change', {
                            bubbles: true,
                            cancelable: true
                        }));
                    }
                }
            })

            const sousLotSelect = document.createElement("select");
            sousLotSelect.className = "sousLotSelect";
            sousLotTd.appendChild(sousLotSelect)

            const blocDeBaseSelect = document.createElement("select");
            blocDeBaseSelect.className = "blocDeBaseSelect";
            blocDeBaseTd.appendChild(blocDeBaseSelect)

            sousLotSelect.addEventListener("change", () => {
                blocDeBaseSelect.innerHTML = '';
                blocDeBaseSelect.innerHTML = '<option value="">Sélectionner un bloc de base</option>';
                this.sousLot = sousLotSelect.value;

                if (this.sousLot) {
                    const lotData = this.lotsData.find(lot => lot.lot === this.lot);
                    const sousElement = lotData.sous_elements.find(se => se.type === this.sousLot);

                    if (sousElement && sousElement.elements) {
                        sousElement.elements.forEach(element => {
                            const option = document.createElement('option');
                            option.value = element;
                            option.textContent = element;
                            blocDeBaseSelect.appendChild(option);
                        });

                        if (this.blocDeBase) {
                            if (sousElement.elements.includes(this.blocDeBase)) {
                                blocDeBaseSelect.value = this.blocDeBase;
                                blocDeBaseSelect.dispatchEvent(new Event('change', {
                                    bubbles: true,
                                    cancelable: true
                                }));
                            }
                        }
                    }
                }
            });

            blocDeBaseSelect.addEventListener("change", () => {
                this.blocDeBase = blocDeBaseSelect.value;
            });
        }

        const specificiteTd = document.createElement("td")
        if (this.specificite === "" || this.specificite) {
            const specificiteTextarea = document.createElement("textarea")
            specificiteTextarea.value = this.specificite
            specificiteTd.appendChild(specificiteTextarea)
        }

        const quantiteBoolTd = document.createElement("td")
        if (this.quantiteBool) {
            const quantiteBoolInput = document.createElement("input")
            quantiteBoolInput.type = "checkbox"
            quantiteBoolInput.checked = true
            if (this.quantiteBool == "false") {
                quantiteBoolInput.checked = false
            }
            quantiteBoolTd.appendChild(quantiteBoolInput)
        }

        const descriptionTd = document.createElement("td")
        if (this.description) {
            descriptionTd.textContent = this.description
        }

        const uniteTd = document.createElement("td")
        if (this.unite) {
            uniteTd.textContent = this.unite
        }

        const quantiteTd = document.createElement("td")
        if (this.quantite) {
            quantiteTd.textContent = this.quantite
        }

        const totalTd = document.createElement("td")
        if (this.montantTotal) {
            totalTd.textContent = this.montantTotal
        }

        tr.append(lotTd, sousLotTd, blocDeBaseTd, specificiteTd, quantiteBoolTd, descriptionTd, uniteTd, quantiteTd, totalTd)

        this.tbody.appendChild(tr)
    }
}