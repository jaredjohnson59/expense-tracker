import Papa from 'papaparse';
    // XLSX is loaded via CDN in index.html, access it via window.XLSX

    // DOM Elements (same as before)
    const fileInput = document.getElementById('csvFile');
    const tableHead = document.querySelector('#dataTable thead');
    const tableBody = document.querySelector('#dataTable tbody');
    const newTagInput = document.getElementById('newTagInput');
    const addTagButton = document.getElementById('addTagButton');
    const availableTagsContainer = document.getElementById('availableTagsDisplay');
    const exportTagSelect = document.getElementById('exportTagSelect');
    const exportButton = document.getElementById('exportButton');
    const importButton = document.getElementById('importButton');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const selectAllCheckbox = document.getElementById('selectAllCheckbox');
    const bulkTagSelect = document.getElementById('bulkTagSelect');
    const bulkTagButton = document.getElementById('bulkTagButton');

    // State (same as before)
    let parsedData = [];
    let tableHeaders = [];
    let availableTags = ["Business", "Personal", "Important"];
    let rowTags = {};
    let currentSort = { column: null, direction: 'asc' };
    const amountColumnName = 'Amount';

    // --- Tag Management --- (Functions remain the same)
    function renderAvailableTags() {
      availableTagsContainer.innerHTML = '<strong>Available Tags:</strong> ';
      bulkTagSelect.innerHTML = '<option value="">Assign Tag to Selected...</option>';
      while (exportTagSelect.options.length > 1) {
          exportTagSelect.remove(1);
      }
      availableTags.forEach(tag => {
        const tagElement = document.createElement('span');
        tagElement.classList.add('tag');
        if (!["Business", "Personal", "Important"].includes(tag)) {
          tagElement.classList.add('custom');
        }
        tagElement.textContent = tag;
        availableTagsContainer.appendChild(tagElement);
        const option = document.createElement('option');
        option.value = tag;
        option.textContent = tag;
        exportTagSelect.appendChild(option.cloneNode(true));
        bulkTagSelect.appendChild(option.cloneNode(true));
      });
      populateExportTagDropdown();
    }

    function populateExportTagDropdown() {
        const currentSelection = exportTagSelect.value;
        if (availableTags.includes(currentSelection)) {
            exportTagSelect.value = currentSelection;
        } else if (exportTagSelect.options.length > 1) {
            exportTagSelect.value = "";
        }
    }

    function addTag() {
      const newTagName = newTagInput.value.trim();
      if (newTagName && !availableTags.includes(newTagName)) {
        availableTags.push(newTagName);
        newTagInput.value = '';
        renderAvailableTags();
      } else if (!newTagName) {
        alert("Please enter a tag name.");
      } else {
        alert(`Tag "${newTagName}" already exists.`);
      }
    }

    // --- Sorting --- (Function remains the same)
    function sortTable(columnIndex) {
        if (columnIndex < 1) return;
        const columnKey = tableHeaders[columnIndex - 1];
        const isNumeric = columnKey.toLowerCase() === amountColumnName.toLowerCase();

        if (currentSort.column === columnKey) {
            currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
        } else {
            currentSort.column = columnKey;
            currentSort.direction = 'asc';
        }

        parsedData.sort((a, b) => {
            const keyA = Object.keys(a).find(k => k.toLowerCase() === columnKey.toLowerCase());
            const keyB = Object.keys(b).find(k => k.toLowerCase() === columnKey.toLowerCase());
            const valA = keyA ? a[keyA] : null;
            const valB = keyB ? b[keyB] : null;
            let comparison = 0;

            if (isNumeric) {
                const numA = (valA !== null && !isNaN(parseFloat(valA))) ? parseFloat(valA) : -Infinity;
                const numB = (valB !== null && !isNaN(parseFloat(valB))) ? parseFloat(valB) : -Infinity;
                comparison = numA - numB;
            } else {
                const strA = String(valA || '').toLowerCase();
                const strB = String(valB || '').toLowerCase();
                if (strA < strB) comparison = -1;
                else if (strA > strB) comparison = 1;
            }
            return currentSort.direction === 'asc' ? comparison : comparison * -1;
        });
        displayTable(parsedData); // Re-render
    }

    // --- CSV/XLSX Parsing and Table Display --- (Functions updated/added)

    // Generic function to parse a single file (CSV or XLSX)
    function parseFile(file) {
        return new Promise((resolve, reject) => {
            const fileName = file.name.toLowerCase();
            const reader = new FileReader();

            reader.onload = (event) => {
                try {
                    const fileData = event.target.result;
                    let data = [];
                    let headers = [];

                    if (fileName.endsWith('.csv')) {
                        // Use PapaParse for CSV
                        const result = Papa.parse(fileData, {
                            header: true,
                            skipEmptyLines: true
                        });
                        if (result.errors.length > 0) {
                            console.error(`PapaParse errors in ${file.name}:`, result.errors);
                            // Decide how to handle errors, maybe reject or resolve with partial data/error info
                            // reject(`Parsing errors in CSV ${file.name}`); return;
                        }
                        data = result.data;
                        headers = result.meta.fields || (data.length > 0 ? Object.keys(data[0]) : []);

                    } else if (fileName.endsWith('.xlsx')) {
                        // Use SheetJS (XLSX) for Excel files
                        const workbook = window.XLSX.read(fileData, { type: 'binary' }); // Use window.XLSX
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];
                        data = window.XLSX.utils.sheet_to_json(worksheet, { raw: false }); // raw: false attempts date/number conversion
                        // Extract headers from the data if available
                        if (data.length > 0) {
                            headers = Object.keys(data[0]);
                        }
                    } else {
                        reject(`Unsupported file type: ${file.name}`);
                        return;
                    }
                    resolve({ data, headers });
                } catch (error) {
                    console.error(`Error processing file ${file.name}:`, error);
                    reject(`Failed to process file ${file.name}: ${error.message}`);
                }
            };

            reader.onerror = (event) => {
                reject(`File could not be read: ${file.name}`);
            };

            // Read file based on type
            if (fileName.endsWith('.csv')) {
                reader.readAsText(file); // PapaParse needs text
            } else if (fileName.endsWith('.xlsx')) {
                reader.readAsBinaryString(file); // XLSX needs binary string or array buffer
            } else {
                 reject(`Unsupported file type: ${file.name}`); // Should not happen if 'accept' works
            }
        });
    }


    function displayTable(data) {
      // Render Header Row
      if (tableHead.children.length === 0 || tableHead.dataset.headers !== JSON.stringify(tableHeaders)) {
          tableHead.innerHTML = '';
          const headerRow = document.createElement('tr');
          const thCheckbox = document.createElement('th');
          headerRow.appendChild(thCheckbox);

          tableHeaders.forEach((header, index) => {
            const th = document.createElement('th');
            th.textContent = header;
            th.dataset.columnIndex = index + 1;
            th.addEventListener('click', () => sortTable(index + 1));
            if (currentSort.column === header) {
                const arrow = document.createElement('span');
                arrow.classList.add('sort-arrow');
                arrow.textContent = currentSort.direction === 'asc' ? ' ▲' : ' ▼';
                th.appendChild(arrow);
            }
            headerRow.appendChild(th);
          });

          const thTags = document.createElement('th');
          thTags.textContent = 'Tags';
          headerRow.appendChild(thTags);
          tableHead.appendChild(headerRow);
          tableHead.dataset.headers = JSON.stringify(tableHeaders);
      } else {
          // Update sort indicators only
          const thElements = tableHead.querySelectorAll('th[data-column-index]');
          thElements.forEach(th => {
              const headerText = th.textContent.replace(/ [▲▼]$/, '').trim();
              const existingArrow = th.querySelector('.sort-arrow');
              if (existingArrow) existingArrow.remove();
              if (currentSort.column === headerText) {
                  const arrow = document.createElement('span');
                  arrow.classList.add('sort-arrow');
                  arrow.textContent = currentSort.direction === 'asc' ? ' ▲' : ' ▼';
                  th.appendChild(arrow);
              }
          });
      }

      // Render Data Rows
      tableBody.innerHTML = '';
      selectAllCheckbox.checked = false;

      data.forEach((row) => {
        const originalIndex = row._originalIndex;
        const dataRow = document.createElement('tr');
        dataRow.dataset.originalIndex = originalIndex;

        const tdCheckbox = document.createElement('td');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.classList.add('row-checkbox');
        checkbox.dataset.originalIndex = originalIndex;
        tdCheckbox.appendChild(checkbox);
        dataRow.appendChild(tdCheckbox);

        tableHeaders.forEach(header => {
          const td = document.createElement('td');
          // Find the key case-insensitively, as XLSX headers might have different casing
          const actualDataKey = Object.keys(row).find(key => key.toLowerCase() === header.toLowerCase());
          const value = actualDataKey ? row[actualDataKey] : '';

          if (header.toLowerCase() === amountColumnName.toLowerCase()) {
              const numericValue = parseFloat(value);
              td.textContent = !isNaN(numericValue) ? numericValue.toFixed(2) : (value || '');
          } else {
              // Handle potential date objects from XLSX parsing if needed
              // if (value instanceof Date) { td.textContent = value.toLocaleDateString(); } else ...
              td.textContent = value || '';
          }
          dataRow.appendChild(td);
        });

        const tdTags = document.createElement('td');
        tdTags.classList.add('tags-cell');
        tdTags.dataset.originalIndex = originalIndex;
        renderRowTags(tdTags, originalIndex);
        dataRow.appendChild(tdTags);

        tableBody.appendChild(dataRow);
      });
    }

    function renderRowTags(cell, originalIndex) { // Use originalIndex
        cell.innerHTML = '';
        const tagsContainer = document.createElement('div');
        tagsContainer.classList.add('tags-container');
        const assigned = rowTags[originalIndex] || [];

        assigned.forEach(tag => {
            const tagElement = document.createElement('span');
            tagElement.classList.add('tag', 'assigned');
             if (!["Business", "Personal", "Important"].includes(tag)) {
                tagElement.classList.add('custom');
            }
            tagElement.textContent = tag;
            tagElement.style.cursor = 'pointer';
            tagElement.title = `Click to remove tag "${tag}"`;
            tagElement.onclick = (e) => {
                e.stopPropagation();
                removeTagFromRow(originalIndex, tag); // Use original index
            };
            tagsContainer.appendChild(tagElement);
        });
        cell.appendChild(tagsContainer);
    }

    // Assign tag using original index
    function assignTagToRow(originalIndex, tag) {
        if (!rowTags[originalIndex]) {
            rowTags[originalIndex] = [];
        }
        if (!rowTags[originalIndex].includes(tag)) {
            rowTags[originalIndex].push(tag);
            const cell = document.querySelector(`.tags-cell[data-original-index="${originalIndex}"]`);
            if (cell) {
                renderRowTags(cell, originalIndex);
            }
        }
    }

    // Remove tag using original index
     function removeTagFromRow(originalIndex, tagToRemove) {
        if (rowTags[originalIndex]) {
            rowTags[originalIndex] = rowTags[originalIndex].filter(tag => tag !== tagToRemove);
            const cell = document.querySelector(`.tags-cell[data-original-index="${originalIndex}"]`);
            if (cell) {
                renderRowTags(cell, originalIndex);
            }
        }
    }

    // Process data: Make amounts positive, ensure numbers, add original index
    function processParsedData(data, startIndex = 0) {
        return data.map((row, index) => {
            const processedRow = { ...row };
            processedRow._originalIndex = startIndex + index;

            const actualAmountKey = Object.keys(processedRow).find(key => key.toLowerCase() === amountColumnName.toLowerCase());
            if (actualAmountKey && processedRow[actualAmountKey] !== null && processedRow[actualAmountKey] !== undefined) {
                // Handle potential currency symbols or commas before parsing
                const cleanedValue = String(processedRow[actualAmountKey]).replace(/[^0-9.-]+/g,"");
                let numericValue = parseFloat(cleanedValue);
                if (!isNaN(numericValue)) {
                    processedRow[actualAmountKey] = Math.abs(numericValue);
                } else {
                     // If cleaning results in non-numeric, keep original (or set to null/0?)
                     // processedRow[actualAmountKey] = processedRow[actualAmountKey]; // Keep original string if parse fails
                }
            }
            return processedRow;
        });
    }

    // --- File Handling --- (Updated to use parseFile)
    async function handleFileSelect(event) {
        const files = event.target.files;
        if (!files || files.length === 0) {
            alert("No files selected.");
            return;
        }

        loadingIndicator.style.display = 'block';
        importButton.disabled = true;

        const currentDataLength = parsedData.length;
        let fileParsePromises = [];

        for (let i = 0; i < files.length; i++) {
            fileParsePromises.push(parseFile(files[i])); // Use the new parseFile function
        }

        try {
            const resultsArray = await Promise.allSettled(fileParsePromises); // Use allSettled to handle individual file errors

            let newData = [];
            let firstHeaders = null;
            let fileErrors = [];

            resultsArray.forEach((result, index) => {
                if (result.status === 'fulfilled') {
                    const { data, headers } = result.value;
                    if (data.length > 0) {
                        // Use headers from the first successfully parsed file
                        if (!firstHeaders && headers && headers.length > 0) {
                            firstHeaders = headers;
                        }
                        // Process data with correct starting index
                        const processed = processParsedData(data, currentDataLength + newData.length);
                        newData = newData.concat(processed);
                    } else {
                         console.warn(`File ${files[index].name} contained no data.`);
                    }
                } else {
                    // Handle rejected promises (parsing errors)
                    console.error(`Failed to parse file ${files[index].name}:`, result.reason);
                    fileErrors.push(`Failed to parse ${files[index].name}: ${result.reason}`);
                }
            });

            if (fileErrors.length > 0) {
                alert("Some files could not be processed:\n- " + fileErrors.join("\n- "));
            }

            if (newData.length > 0) {
                parsedData = parsedData.concat(newData); // Append new data

                // Set/Validate table headers
                if (tableHeaders.length === 0 && firstHeaders) {
                    tableHeaders = firstHeaders;
                } else if (tableHeaders.length > 0 && firstHeaders && JSON.stringify(tableHeaders.map(h=>h.toLowerCase())) !== JSON.stringify(firstHeaders.map(h=>h.toLowerCase()))) {
                    console.warn("Header mismatch detected between files. Using original headers.", tableHeaders, firstHeaders);
                    alert("Warning: Headers in newly imported files differ from existing data. Display might be inconsistent. Using original headers.");
                }

                console.log("Files processed. Combined data:", parsedData);
                currentSort = { column: null, direction: 'asc' }; // Reset sort
                displayTable(parsedData);
            } else if (fileErrors.length === 0) {
                 alert("No new data found in the selected files.");
            }


        } catch (error) { // Catch errors not related to individual file parsing
            console.error("Error processing files:", error);
            alert(`An unexpected error occurred during file processing: ${error}`);
        } finally {
            loadingIndicator.style.display = 'none';
            importButton.disabled = false;
            fileInput.value = ''; // Reset file input
        }
    }


    // --- Export Functionality --- (Remains the same, exports as CSV)
    function exportFilteredData() {
        const selectedTag = exportTagSelect.value;
        if (!selectedTag) {
            alert("Please select a tag to filter by for export.");
            return;
        }
        if (parsedData.length === 0) {
             alert("No data available to export. Please upload a file first.");
            return;
        }

        const filteredRawData = parsedData.filter(row => {
            return rowTags[row._originalIndex] && rowTags[row._originalIndex].includes(selectedTag);
        });

        if (filteredRawData.length === 0) {
            alert(`No transactions found with the tag "${selectedTag}".`);
            return;
        }

        const dataToExport = filteredRawData.map(row => {
             const exportRow = { ...row };
             delete exportRow._originalIndex;
             const actualAmountKey = Object.keys(exportRow).find(key => key.toLowerCase() === amountColumnName.toLowerCase());
             if (actualAmountKey) {
                 const numericValue = parseFloat(exportRow[actualAmountKey]);
                 if (!isNaN(numericValue)) {
                     exportRow[actualAmountKey] = numericValue.toFixed(2);
                 }
             }
             return exportRow;
        });

        const columnsToExport = tableHeaders.filter(h => h !== '_originalIndex');
        const csvString = Papa.unparse(dataToExport, {
            header: true,
            columns: columnsToExport
        });

        const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement("a");
        const url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        const safeTagName = selectedTag.replace(/[^a-z0-9]/gi, '_').toLowerCase();
        link.setAttribute("download", `transactions_${safeTagName}_export.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // --- Bulk Tagging --- (Remains the same)
    function applyBulkTag() {
        const selectedTag = bulkTagSelect.value;
        if (!selectedTag) {
            alert("Please select a tag from the dropdown to apply.");
            return;
        }
        const selectedCheckboxes = tableBody.querySelectorAll('.row-checkbox:checked');
        if (selectedCheckboxes.length === 0) {
            alert("Please select at least one row using the checkboxes.");
            return;
        }
        selectedCheckboxes.forEach(checkbox => {
            const originalIndex = parseInt(checkbox.dataset.originalIndex, 10);
            if (!isNaN(originalIndex)) {
                assignTagToRow(originalIndex, selectedTag);
            }
        });
        bulkTagSelect.value = "";
    }

    // --- Event Listeners --- (Remain the same)
    importButton.addEventListener('click', () => fileInput.click() );
    fileInput.addEventListener('change', handleFileSelect);
    addTagButton.addEventListener('click', addTag);
    newTagInput.addEventListener('keypress', (event) => { if (event.key === 'Enter') addTag(); });
    exportButton.addEventListener('click', exportFilteredData);
    selectAllCheckbox.addEventListener('change', (event) => {
        const isChecked = event.target.checked;
        tableBody.querySelectorAll('.row-checkbox').forEach(checkbox => checkbox.checked = isChecked);
    });
    bulkTagButton.addEventListener('click', applyBulkTag);

    // --- Initial Render ---
    renderAvailableTags();
    displayTable([]);
