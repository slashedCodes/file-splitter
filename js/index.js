// Process the user-submitted zip file.
async function submit() {
    try {
        let file = document.getElementById("file").files[0];
        let reader = new FileReader();

        reader.onload = await async function (event) {
            const zip = await JSZip.loadAsync(event.target.result);
            const filenames = Object.keys(zip.files)
            let headers = []
            let newFilenames = []

            // Get filenames
            await Promise.all(filenames.map(async filename => {
                const blob = await zip.file(filename).async("blob");
                const header = await getHeader(blob);
                headers.push(header + ".docx");
            }));

            // Duplicate check and process array names
            newFilenames = filterArray(headers);

            // Rename files in zip
            for (let i = 0; i < newFilenames.length; i++) {
                const newFilename = newFilenames[i]
                const oldFilename = filenames[i]
                await renameFile(oldFilename, newFilename, zip);
            }

            const finalZip = await zip.generateAsync({ type: "blob" })
            saveAs(finalZip, "renamed.zip")
        }

        reader.readAsArrayBuffer(file);
    } catch (err) {
        console.error("[Error] submit(): ", err);
    }
}

// Rename a file in the zip file.
async function renameFile(oldFilename, newFilename, zip) {
    try {
        const old = await zip.file(oldFilename).async("blob");
        zip.remove(oldFilename);
        zip.file(newFilename, old);
    } catch (err) {
        console.error("[Error] renameFile(): ", err);
    }
}

// Filter filename array.
function filterArray(array) {
    try {
        // Handle duplicates
        let filteredFilenames = []
        filteredFilenames = arrayDuplicateCheck(array);

        // Make sure the filenames are valid
        let processedFilenames = []
        filteredFilenames.forEach(filename => {
            const processedFilename = processFilename(filename);
            processedFilenames.push(processedFilename);
        });

        return processedFilenames;
    } catch (err) {
        console.error("[Error] filterArray(): ", err);
    }
}

// Filter an array of duplicates.
function arrayDuplicateCheck(array) {
    try {
        const counts = {};

        array.forEach(element => {
            counts[element] = (counts[element] || 0) + 1;
        });

        array.forEach((element, index) => {
            if (counts[element] > 1) {
                const uniqueName = generateUniqueName(element, array.slice(0, index));
                array[index] = uniqueName;
                counts[uniqueName] = (counts[uniqueName] || 0) + 1;
            }
        });

        return array;
    } catch (err) {
        console.error("[Error] arrayDuplicateCheck(): ", err);
    }
}

// Generate a unique filename
function generateUniqueName(name, existingNames) {
    try {
        let uniqueName = name;
        let count = 1;

        while (existingNames.includes(uniqueName)) {
            uniqueName = `${name} (${count})`;
            count++;
        }

        return uniqueName;
    } catch (err) {
        console.error("[Error] generateUniqueName(): ", err);
    }
}

// Manipulate a string such as it can be saved as a filename in NTFS.
function processFilename(str) {
    try {
        let newStr = str;

        // List can be expanded.
        newStr = newStr.replaceAll("/", "-");
        newStr = newStr.replaceAll("\\", "-");
        newStr = newStr.replaceAll("|", "-");
        newStr = newStr.replaceAll('"', "'");

        return newStr;
    } catch (err) {
        console.error("[Error] processFilename(): ", err);
    }
}

// Helper function to process xml.
function stringToXml(str) {
    try {
        if (str.charCodeAt(0) === 65279) {
            // BOM sequence
            str = str.substr(1);
        }
        return new DOMParser().parseFromString(str, "text/xml");
    } catch (err) {
        console.error("[Error] stringToXml(): ", err);
    }
}

// Get the first line in a DOCX blob
// Shamelessly stolen and modified from https://docxtemplater.com/faq/#how-can-i-retrieve-the-docx-content-as-text (and it doesn't even use docxtemplater)
async function getHeader(blob) {
    try {
        const zip = await JSZip.loadAsync(blob);
        const documentContent = await zip.file("word/document.xml").async("string");

        // Parse doc
        const xmlDoc = stringToXml(documentContent);
        const paragraphsXml = xmlDoc.getElementsByTagName("w:p");

        // Extract text from the first paragraph
        const firstParagraphXml = paragraphsXml[0];
        const textNodes = firstParagraphXml.getElementsByTagName("w:t");
        let headerText = "";

        // Concatenate text from all text nodes in the first paragraph
        for (let i = 0; i < textNodes.length; i++) {
            const textNode = textNodes[i];
            if (textNode.childNodes && textNode.childNodes.length > 0) {
                headerText += textNode.childNodes[0].nodeValue;
            }
        }

        return headerText;
    } catch (error) {
        console.error("[Error] getHeader(): ", error);
        return null;
    }
}