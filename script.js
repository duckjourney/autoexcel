function processFile() {
  const fileInput = document.getElementById("file-input");
  const startTimeInput = document.getElementById("start-time");
  const endTimeInput = document.getElementById("end-time");
  const file = fileInput.files[0];

  if (file && startTimeInput.value && endTimeInput.value) {
    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const dataRows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
      });

      // Find the index of the "closed_time" column
      const headerRow = dataRows[0];
      let closedTimeColumnIndex = null;
      for (let i = 0; i < headerRow.length; i++) {
        if (headerRow[i] === "closed_time") {
          closedTimeColumnIndex = i;
          break;
        }
      }

      // Ensure that the "closed_time" column was found
      if (closedTimeColumnIndex === null) {
        console.error("Could not find a column named 'closed_time'.");
        return;
      }

      // Parse start and end times
      const [startHours, startMinutes] = startTimeInput.value.split(":");
      const [endHours, endMinutes] = endTimeInput.value.split(":");
      const startTotalMinutes =
        parseInt(startHours) * 60 + parseInt(startMinutes);
      const endTotalMinutes = parseInt(endHours) * 60 + parseInt(endMinutes);

      // Calculate the number of 15-minute intervals
      const intervalCount = (endTotalMinutes - startTotalMinutes) / 15;
      const counters = new Array(intervalCount).fill(0);

      // Iterate over each row in the Excel file
      for (let i = 1; i < dataRows.length; i++) {
        const row = dataRows[i];
        if (row.length > 0) {
          const dateTimeString = row[closedTimeColumnIndex];
          const [dateString, timeString] = dateTimeString.split(" ");
          const [hours, minutes] = timeString.split(":");
          const minutesTotal = parseInt(hours) * 60 + parseInt(minutes);

          // Count items based on time intervals
          if (
            minutesTotal >= startTotalMinutes &&
            minutesTotal < endTotalMinutes
          ) {
            const intervalIndex = Math.floor(
              (minutesTotal - startTotalMinutes) / 15
            );
            counters[intervalIndex]++;
          }
        }
      }

      // Print the results
      const resultsTbody = document.getElementById("results-tbody");
      for (let i = 0; i < counters.length; i++) {
        const intervalStart = startTotalMinutes + i * 15;
        const intervalEnd = intervalStart + 15;
        const intervalStartStr = `${String(
          Math.floor(intervalStart / 60)
        ).padStart(2, "0")}:${String(intervalStart % 60).padStart(2, "0")}`;
        const intervalEndStr = `${String(Math.floor(intervalEnd / 60)).padStart(
          2,
          "0"
        )}:${String(intervalEnd % 60).padStart(2, "0")}`;

        const row = document.createElement("tr");
        const intervalCell = document.createElement("td");
        intervalCell.textContent = `${intervalStartStr}-${intervalEndStr}`;
        row.appendChild(intervalCell);

        const countCell = document.createElement("td");
        countCell.textContent = counters[i];
        row.appendChild(countCell);

        resultsTbody.appendChild(row);
      }
    };
    reader.readAsArrayBuffer(file);
  }
}

function copyTableToClipboard() {
  const table = document.getElementById('results-table');
  const range = document.createRange();
  range.selectNode(table);
  window.getSelection().removeAllRanges();
  window.getSelection().addRange(range);
  document.execCommand('copy');
  window.getSelection().removeAllRanges();
}
