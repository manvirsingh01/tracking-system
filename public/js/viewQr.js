document.getElementById('viewLogButton').addEventListener('click', async () => {
    const logTable = document.getElementById('logTable').querySelector('tbody');
    const logDataDiv = document.getElementById('logData');

    try {
        // Fetch the log data from the server
        const response = await fetch(`/getLog/<%= documentDetails.id %>`);
        if (response.ok) {
            const logData = await response.json();
            logTable.innerHTML = ''; // Clear any existing rows

            // Populate the table with log data
            logData.forEach(entry => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${entry.Date}</td>
                    <td>${entry.InTime}</td>
                    <td>${entry.Place}</td>
                    <td>${entry.OutTime}</td>
                `;
                logTable.appendChild(row);
            });

            logDataDiv.style.display = 'block'; // Show the log section
        } else {
            alert('Error fetching log data.');
        }
    } catch (error) {
        console.error('Error:', error);
        alert('An error occurred while fetching log data.');
    }
});