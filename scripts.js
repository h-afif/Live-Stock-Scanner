document.getElementById('stock-form').addEventListener('submit', function(event) {
    event.preventDefault();
    const symbol = document.getElementById('symbol').value.toUpperCase();
    const startDate = document.getElementById('start-date').value;
    const endDate = document.getElementById('end-date').value;
    const apiKey = 'Itewa654mtL2zbYxg5l78qfRMiLa0cyq'; // Polygon.io API key

    // Define URLs for real-time and company profile data
    const realTimeUrl = `https://api.polygon.io/v2/aggs/ticker/${symbol}/prev?apiKey=${apiKey}`;
    const profileUrl = `https://api.polygon.io/v1/meta/symbols/${symbol}/company?apiKey=${apiKey}`;

    // Fetch both real-time data and company profile data
    Promise.all([
        fetch(realTimeUrl).then(response => response.json()),
        fetch(profileUrl).then(response => response.json())
    ])
    .then(([realTimeData, profileData]) => {
        console.log('Real-time Data:', realTimeData); // Log real-time data
        console.log('Profile Data:', profileData); // Log profile data

        if (!realTimeData || !realTimeData.results) {
            document.getElementById('results').innerHTML = '<p>No real-time data found or invalid symbol.</p>';
            return;
        }

        // Extract data from profile data
        const marketCap = profileData && profileData.marketCap; // Adjust field name based on API response
        const sector = profileData && profileData.sector ? profileData.sector : 'N/A';

        const metrics = {
            'Symbol': symbol,
            'Price': realTimeData.results[0].c ? `$${realTimeData.results[0].c.toFixed(2)}` : 'N/A',
            'Market Cap': marketCap ? `$${(marketCap / 1e9).toFixed(2)} Billion` : 'N/A', // Adjust format as needed
            'Volume': realTimeData.results[0].v ? realTimeData.results[0].v.toLocaleString() : 'N/A',
            'Change %': realTimeData.results[0].c && realTimeData.results[0].o ? `${(((realTimeData.results[0].c - realTimeData.results[0].o) / realTimeData.results[0].o) * 100).toFixed(2)}%` : 'N/A',
            'Previous Close': realTimeData.results[0].c ? `$${realTimeData.results[0].c.toFixed(2)}` : 'N/A',
            'Sector': sector
        };

        // Generate HTML for metrics
        let tableHTML = '<h2>Metrics</h2>';
        tableHTML += '<table class="table">';
        tableHTML += '<thead><tr class="header-row">';
        tableHTML += '<th>Metric</th><th>Value</th>';
        tableHTML += '</tr></thead>';
        tableHTML += '<tbody>';
        tableHTML += Object.entries(metrics).map(([key, value]) => `<tr class="data-row"><td>${key}</td><td>${value}</td></tr>`).join('');
        tableHTML += '</tbody>';
        tableHTML += '</table>';

        // Fetch historical data
        const historicalUrl = `https://api.polygon.io/v2/aggs/ticker/${symbol}/range/1/day/${startDate}/${endDate}?apiKey=${apiKey}`;
        fetch(historicalUrl)
            .then(response => response.json())
            .then(historicalData => {
                console.log('Historical Data:', historicalData); // Log historical data

                let historicalHTML = '<h2>Historical Data</h2>';
                if (historicalData.results && historicalData.results.length > 0) {
                    historicalHTML += '<table class="table">';
                    historicalHTML += '<thead><tr class="header-row">';
                    historicalHTML += '<th>Date</th><th>Open</th><th>Close</th><th>Volume</th>';
                    historicalHTML += '</tr></thead>';
                    historicalHTML += '<tbody>';
                    historicalHTML += historicalData.results.map(data => 
                        `<tr class="data-row">
                            <td>${new Date(data.t).toLocaleDateString()}</td>
                            <td>$${data.o.toFixed(2)}</td>
                            <td>$${data.c.toFixed(2)}</td>
                            <td>${data.v.toLocaleString()}</td>
                        </tr>`).join('');
                    historicalHTML += '</tbody>';
                    historicalHTML += '</table>';
                } else {
                    historicalHTML += '<p>No historical data available</p>';
                }

                document.getElementById('results').innerHTML = tableHTML + historicalHTML;

                // Create download button
                const downloadBtn = document.createElement('button');
                downloadBtn.id = 'download-excel';
                downloadBtn.textContent = 'Download Excel';
                document.getElementById('results').appendChild(downloadBtn);

                // Handle download
                downloadBtn.addEventListener('click', function() {
                    const wb = XLSX.utils.book_new();

                    // Add metrics sheet
                    const metricsWS = XLSX.utils.json_to_sheet(Object.entries(metrics).map(([key, value]) => ({ Metric: key, Value: value })));
                    XLSX.utils.book_append_sheet(wb, metricsWS, 'Metrics');

                    // Add historical data sheet
                    const historicalWS = XLSX.utils.json_to_sheet(historicalData.results.map(data => ({
                        Date: new Date(data.t).toLocaleDateString(),
                        Open: data.o.toFixed(2),
                        Close: data.c.toFixed(2),
                        Volume: data.v.toLocaleString()
                    })));
                    XLSX.utils.book_append_sheet(wb, historicalWS, 'Historical Data');

                    // Write to file
                    XLSX.writeFile(wb, `${symbol}_data.xlsx`);
                });
            })
            .catch(error => {
                console.error('Error fetching historical data:', error);
                document.getElementById('results').innerHTML = '<p>Error fetching historical data. Please try again.</p>';
            });
    })
    .catch(error => {
        console.error('Error fetching data:', error);
        document.getElementById('results').innerHTML = '<p>Error fetching data. Please try again.</p>';
    });
});
