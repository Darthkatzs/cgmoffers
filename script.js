class QuotationGenerator {
    constructor() {
        this.initializeEventListeners();
        this.updateTotals();
    }

    initializeEventListeners() {
        // Add row buttons
        document.getElementById('addOneTimeRow').addEventListener('click', () => {
            this.addCostRow('oneTimeCosts');
        });
        
        document.getElementById('addRecurringRow').addEventListener('click', () => {
            this.addCostRow('recurringCosts');
        });

        // Generate PDF button
        document.getElementById('generatePDF').addEventListener('click', () => {
            this.generatePDF();
        });

        // Initial event listeners for existing rows
        this.attachRowEventListeners();
    }

    addCostRow(containerId) {
        const container = document.getElementById(containerId);
        const newRow = document.createElement('div');
        newRow.className = 'cost-row';
        
        newRow.innerHTML = `
            <input type="text" class="material-name" placeholder="Materiaal/Service naam">
            <input type="number" class="quantity" placeholder="1" min="1" value="1">
            <input type="number" class="unit-price" placeholder="0.00" min="0" step="0.01">
            <input type="number" class="total-price" readonly>
            <button type="button" class="remove-row">×</button>
        `;
        
        container.appendChild(newRow);
        this.attachRowEventListeners();
    }

    attachRowEventListeners() {
        // Remove row buttons
        document.querySelectorAll('.remove-row').forEach(button => {
            button.onclick = (e) => {
                const row = e.target.closest('.cost-row');
                row.remove();
                this.updateTotals();
            };
        });

        // Quantity and price input listeners
        document.querySelectorAll('.quantity, .unit-price').forEach(input => {
            input.addEventListener('input', (e) => {
                this.updateRowTotal(e.target.closest('.cost-row'));
                this.updateTotals();
            });
        });
    }

    updateRowTotal(row) {
        const quantity = parseFloat(row.querySelector('.quantity').value) || 0;
        const unitPrice = parseFloat(row.querySelector('.unit-price').value) || 0;
        const total = quantity * unitPrice;
        row.querySelector('.total-price').value = total.toFixed(2);
    }

    updateTotals() {
        // Update all row totals first
        document.querySelectorAll('.cost-row').forEach(row => {
            this.updateRowTotal(row);
        });

        // Calculate one-time costs total
        const oneTimeRows = document.querySelectorAll('#oneTimeCosts .cost-row');
        let oneTimeTotal = 0;
        oneTimeRows.forEach(row => {
            const total = parseFloat(row.querySelector('.total-price').value) || 0;
            oneTimeTotal += total;
        });

        // Calculate recurring costs total
        const recurringRows = document.querySelectorAll('#recurringCosts .cost-row');
        let recurringTotal = 0;
        recurringRows.forEach(row => {
            const total = parseFloat(row.querySelector('.total-price').value) || 0;
            recurringTotal += total;
        });

        // Update display
        document.getElementById('oneTimeTotal').textContent = oneTimeTotal.toFixed(2);
        document.getElementById('recurringTotal').textContent = recurringTotal.toFixed(2);
        document.getElementById('finalOneTimeTotal').textContent = oneTimeTotal.toFixed(2);
        document.getElementById('finalRecurringTotal').textContent = recurringTotal.toFixed(2);
    }

    getFormData() {
        const formData = {
            companyName: document.getElementById('companyName').value,
            contactName: document.getElementById('contactName').value,
            address: document.getElementById('address').value,
            postalCode: document.getElementById('postalCode').value,
            city: document.getElementById('city').value,
            companyId: document.getElementById('companyId').value,
            date: new Date().toLocaleDateString('nl-NL'),
            oneTimeCosts: [],
            recurringCosts: []
        };

        // Collect one-time costs
        document.querySelectorAll('#oneTimeCosts .cost-row').forEach(row => {
            const materialName = row.querySelector('.material-name').value;
            const quantity = parseInt(row.querySelector('.quantity').value) || 0;
            const unitPrice = parseFloat(row.querySelector('.unit-price').value) || 0;
            const total = parseFloat(row.querySelector('.total-price').value) || 0;
            
            if (materialName && quantity > 0) {
                formData.oneTimeCosts.push({
                    material: materialName,
                    quantity: quantity,
                    unitPrice: unitPrice,
                    total: total
                });
            }
        });

        // Collect recurring costs
        document.querySelectorAll('#recurringCosts .cost-row').forEach(row => {
            const materialName = row.querySelector('.material-name').value;
            const quantity = parseInt(row.querySelector('.quantity').value) || 0;
            const unitPrice = parseFloat(row.querySelector('.unit-price').value) || 0;
            const total = parseFloat(row.querySelector('.total-price').value) || 0;
            
            if (materialName && quantity > 0) {
                formData.recurringCosts.push({
                    material: materialName,
                    quantity: quantity,
                    unitPrice: unitPrice,
                    total: total
                });
            }
        });

        return formData;
    }

    async generatePDF() {
        try {
            const formData = this.getFormData();
            
            // Validate required fields
            if (!formData.companyName || !formData.contactName || !formData.address) {
                alert('Vul alle verplichte velden in voordat u de PDF genereert.');
                return;
            }

            if (formData.oneTimeCosts.length === 0 && formData.recurringCosts.length === 0) {
                alert('Voeg minimaal één kostenpost toe voordat u de PDF genereert.');
                return;
            }

            // Load the Word template
            const templateResponse = await fetch('standaardofferte Compufit NL.docx');
            if (!templateResponse.ok) {
                throw new Error('Kan sjabloon niet laden');
            }
            
            const templateContent = await templateResponse.arrayBuffer();
            
            // Process the template with docxtemplater
            const zip = new PizZip(templateContent);
            const doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            // Calculate totals
            const oneTimeTotal = formData.oneTimeCosts.reduce((sum, item) => sum + item.total, 0);
            const recurringTotal = formData.recurringCosts.reduce((sum, item) => sum + item.total, 0);

            // Prepare data for template
            const templateData = {
                companyName: formData.companyName,
                contactName: formData.contactName,
                fullAddress: `${formData.address}, ${formData.postalCode} ${formData.city}`,
                companyId: formData.companyId,
                date: formData.date,
                oneTimeCosts: formData.oneTimeCosts,
                recurringCosts: formData.recurringCosts,
                oneTimeTotal: oneTimeTotal.toFixed(2),
                recurringTotal: recurringTotal.toFixed(2),
                hasOneTimeCosts: formData.oneTimeCosts.length > 0,
                hasRecurringCosts: formData.recurringCosts.length > 0
            };

            // Render the document
            doc.render(templateData);

            // Generate the output
            const output = doc.getZip().generate({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });

            // Save the file
            const fileName = `Offerte_${formData.companyName.replace(/[^a-zA-Z0-9]/g, '_')}_${new Date().getTime()}.docx`;
            saveAs(output, fileName);

            alert('Offerte succesvol gegenereerd! Het bestand is gedownload.');

        } catch (error) {
            console.error('Error generating PDF:', error);
            alert('Er is een fout opgetreden bij het genereren van de PDF. Controleer of het sjabloon beschikbaar is en probeer opnieuw.');
        }
    }
}

// Initialize the application when the page loads
document.addEventListener('DOMContentLoaded', () => {
    new QuotationGenerator();
});