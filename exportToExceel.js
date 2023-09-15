const xlsx = require('node-xlsx');

const exportToExceel = async (data) => {
  try {
    // Check if data is empty
    if (!data.length) {
      throw new Error("No data to export.");
    }
    
    // Create a worksheet
    const worksheet = xlsx.build([{ name: 'Data', data: data }]);

    // Directly return the buffer
    return worksheet;
  } catch (error) {
    console.error('Error exporting data to Excel:', error);
    throw error;
  }
};

module.exports = exportToExceel;











// Use this to ONLY push download to USER (ie. not save on server):

// const xlsx = require('node-xlsx');
// const fs = require('fs');
// const path = require('path');
// const util = require('util');

// const writeFile = util.promisify(fs.writeFile);
// const readFile = util.promisify(fs.readFile);
// const unlinkFile = util.promisify(fs.unlink);

// const exportToExceel = async (data) => {
//   try {

//      // Check if data is empty
//      if (!data.length) {
//       throw new Error("No data to export.");
//     }
//     // Prepare data for the worksheet
//     const wsData = data.map(obj => Object.values(obj));

//     // Add headers to wsData
//     wsData.unshift(Object.keys(data[0]));

//     // Create a worksheet
//     const worksheet = xlsx.build([{ name: 'Data', data: wsData }]);

//     // Create a temporary file path
//     const tmpFilePath = path.resolve(__dirname, 'tmpfile.xlsx');

//     // Write workbook to the temporary file
//     await writeFile(tmpFilePath, worksheet, 'binary');

//     // Read the file into a buffer
//     const buffer = await readFile(tmpFilePath);

//     // Delete the temporary file
//     await unlinkFile(tmpFilePath);

//     return buffer;
//   } catch (error) {
//     console.error('Error exporting data to Excel:', error);
//     throw error;
//   }
// };

// module.exports = exportToExceel;