document.addEventListener('DOMContentLoaded', () => {
    const startButton = document.getElementById('start');
    const resultParagraph = document.getElementById('result');
    const fileInput = document.getElementById('fileInput');
    const spreadsheetTable = document.getElementById('spreadsheet');
    const recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();

    recognition.onstart = () => {
        resultParagraph.textContent = 'Listening...';
    };

    recognition.onspeechend = () => {
        recognition.stop();
    };

    recognition.onresult = (event) => {
        const transcript = event.results[0][0].transcript;
        resultParagraph.textContent = `You said: ${transcript}`;

        const numbers = transcript.match(/\d+/g);
        if (numbers) {
            highlightNumbers(numbers);
        }
    };

    startButton.addEventListener('click', () => {
        recognition.start();
    });

    fileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const htmlString = XLSX.utils.sheet_to_html(worksheet);
            spreadsheetTable.innerHTML = htmlString;
        };
        reader.readAsArrayBuffer(file);
    });

    function highlightNumbers(numbers) {
        const rows = spreadsheetTable.querySelectorAll('tr');
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            cells.forEach(cell => {
                numbers.forEach(number => {
                    if (cell.textContent.includes(number)) {
                        cell.style.backgroundColor = 'yellow';
                    }
                });
            });
        });
    }
});

