// File: /jupyter-to-webapp/jupyter-to-webapp/src/static/js/scripts.js

document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('inputForm');
    const outputSection = document.getElementById('outputSection');

    form.addEventListener('submit', function(event) {
        event.preventDefault();
        const formData = new FormData(form);
        const data = {};
        formData.forEach((value, key) => {
            data[key] = value;
        });

        fetch('/process', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(data),
        })
        .then(response => response.json())
        .then(data => {
            outputSection.innerHTML = '';
            const result = document.createElement('div');
            result.innerHTML = `<h3>Results:</h3><pre>${JSON.stringify(data, null, 2)}</pre>`;
            outputSection.appendChild(result);
        })
        .catch((error) => {
            console.error('Error:', error);
        });
    });
});