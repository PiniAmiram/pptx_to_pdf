const { execFile } = require('child_process');

const pptxToPdf = (pptxPath, outputPdf, callback) => {
    const scriptPath = "path/to/your/python_script.py"; // Replace with your script path
    const args = [pptxPath, "output_folder", outputPdf];

    execFile("python", [scriptPath, ...args], (error, stdout, stderr) => {
        if (error) {
            console.error(`Error: ${error.message}`);
            callback(error, null);
            return;
        }
        if (stderr) {
            console.error(`Stderr: ${stderr}`);
        }
        console.log(`Stdout: ${stdout}`);
        callback(null, outputPdf);
    });
};

pptxToPdf("sample.pptx", "output.pdf", (err, result) => {
    if (err) {
        console.error("Conversion failed:", err);
        return;
    }
    console.log("PDF created successfully:", result);
});
