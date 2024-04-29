const express = require('express');
const multer = require('multer');
const path = require('path');
const createReport = require('docx-templates').default;
const fs = require('fs');

const app = express();
const port = 30000;

// Configure multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Serve static files from the 'public' directory
app.use(express.static(path.join(__dirname, 'public')));

// Route to handle file uploads
app.post('/upload', upload.single('file'), async (req, res) => {
 try {
    // Read the uploaded file
    const template = fs.readFileSync(req.file.path);
    // Generate the report with custom command delimiters and the custom 'tile' command
    const buffer = await createReport({
      template,
      data: {
        username: "លន លឹមហៃ",
        age: 22,
        position: "មន្រ្តីកិច្ចសន្យា",
        office: "ក្រសួងប្រៃសណីយ៍",
        department: "គណៈកម្មាធិការរដ្ឋាភិបាលឌីជីថល",
        organization: "អង្គភាពប្រឆាំងអំពើពុករលួយ​",
        phoneNumber: "០៧៧ ៦០៩​ ០៦៤",
        profilePicture: {
          width: 3, // Width in cm
          height: 3, // Height in cm
          data: fs.readFileSync('/home/lim-hai/docx-templates/examples/example-webpack/public/signature.png'), // Image data
          extension: '.jpg', // Image file extension
        },
      },
      cmdDelimiter: ['{', '}'], // Specify custom delimiters here
      commands: {
        tile: (params, context) => {
          // Implement your custom logic here
          // `params` contains the parameters passed to the command
          // `context` provides access to the document and other utilities
          // For demonstration, let's just log the parameters
          console.log('Tile command called with parameters:', params);
          // You can manipulate the document using the `context` object here
        },
      },
    });
    // Write the generated report to a file
    fs.writeFileSync('report.docx', buffer);
    // Send a success response
    res.send('Report generated successfully');
 } catch (error) {
    // Log the error and send a 500 response
    console.error('Error generating report:', error);
    res.status(500).send('Error generating report');
 }
});

// Route to serve the index.html file
app.get('/', (req, res) => {
 res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Start the server
app.listen(port, err => {
 if (err) {
    console.error('something bad happened', err);
    return;
 }
 console.log(`server is listening on ${port}`);
});
