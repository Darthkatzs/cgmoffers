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
            description: document.getElementById('description').value,
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

            // Show processing message
            const button = document.getElementById('generatePDF');
            const originalText = button.textContent;
            button.textContent = 'Bezig met genereren...';
            button.disabled = true;

            try {
                // Send data to the final quotation server for processing
                const response = await fetch('http://localhost:8001/generate-quotation', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify(formData)
                });

                if (!response.ok) {
                    throw new Error(`Server error: ${response.status}`);
                }

                const result = await response.json();
                
                if (result.success) {
                    alert(`Offerte succesvol gegenereerd! Bestand: ${result.filename}`);
                    
                    // Trigger download
                    if (result.download_url) {
                        const downloadLink = document.createElement('a');
                        downloadLink.href = `http://localhost:8001${result.download_url}`;
                        downloadLink.download = result.filename;
                        downloadLink.style.display = 'none';
                        document.body.appendChild(downloadLink);
                        downloadLink.click();
                        document.body.removeChild(downloadLink);
                    }
                    
                } else {
                    throw new Error(result.error || 'Unknown server error');
                }

            } finally {
                // Restore button state
                button.textContent = originalText;
                button.disabled = false;
            }

        } catch (error) {
            console.error('Error generating quotation:', error);
            alert(`Er is een fout opgetreden bij het genereren van de offerte: ${error.message}`);
        }
    }
}

// Initialize the application when the page loads
document.addEventListener('DOMContentLoaded', () => {
    new QuotationGenerator();
});