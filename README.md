# KYC Workbench

Run locally with Python 3.10 or newer and Tk. On Windows, the standard Python installer includes Tk by default. On Linux, install your distribution's Python Tk package if Tk is missing.

```shell
python -m pip install -r requirements.txt
python app.py
```

Choose **Open image** for one scan or **Open folder** for a batch. To combine batches, use **Open multiple folders**, click **Add folder** for each folder, and select **Open folders**. The app uses each folder's configured key and combines all rows into one Excel export. **Download all images** saves all loaded scans as readable JPEG or PNG files in the folder you choose. Images from multiple folders go into numbered subfolders so repeated filenames stay separate. **File > Decrypt Images...** also offers a JPEG-only export. Use an output drive with hard link support, such as NTFS.

## Export records from Client 2 grid pages

1. Open your encrypted JPG folder and check that the page preview is readable.
2. Click **Export Excel**. Workbench recognizes new three-part ID prefixes, such as `FbK_BrT-NaY`, without a code change. It checks the IDs against your selected numbered **KYC MI** workbook and looks for a matching workbook in the scan folder, its parent, or **Downloads** when the selected file has no matches. If no MI workbook matches, Workbench can recover verified rows from compatible earlier exports. When a loaded folder has no verified records, **Export Excel** stops and identifies the source problem before saving a file. This step uses local Tesseract OCR or the installed Windows OCR language.
3. The app reads printed record IDs and uses the matching MI rows to populate the 22 columns. On ruled two-column scans, it reads the ID area of both columns and checks the Center Code against the printed lines. It restores a missing `Y` or `y` only when two OCR views agree on its position and case. Records follow each numbered grid row from left to right. Image Name holds the page number. **Scan review** lists unmatched IDs and flags gaps between detected numbers for comparison with the original pages.

Use **File > Export OCR review...** to save raw OCR text and detected record IDs for troubleshooting. That workbook has **Scan review** and **Detected IDs** sheets.

Recovered workbooks start with `KYC_Recovered_`. In a recovered workbook, or on any scan with no matches in the MI workbook, complete detected IDs keep their original scan number in **KYC Data**, with blank data fields and a **REVIEW** remark. Partial IDs and scans with no readable ID appear in **Scan review**. The export button runs without a source picker. You can set a numbered source explicitly from **File > Change source workbook...**.

For manually entered or labeled forms, use **File > Export reviewed form rows**. If the loaded scans have no entered values, this command runs the grid export.

For a readable XLSX, choose its worksheet and header row, then map the actual columns. The example workbook uses numbered headers, so its mapping needs to be supplied by the operator. Exact named-header matches are preselected. Ensure identifier, postal code and phone columns are stored as text to preserve leading zeros.

The tool applies the client examples for US nationality, gender, name case, dates, email, phone numbers and missing values. Review generated remarks before delivery. The supplied XLSX already contains readable cell values and uses the Excel import workflow.

## Automatic image decryption on Windows

Keep your original `MASTER_DECRYPT_PROMPT.md` in the encrypted image folder. The app reads its `KEY STRING` line when you select **Open image folder**. The app folder, its parent and your Downloads folder are also checked in that order after the image folder. An explicit `KYC_IMAGE_KEY` environment variable overrides file discovery. If no override is provided, the app uses the built-in Client 2 key. **Set legacy key** provides a session-only manual override.

For a whole folder at once:

```powershell
python decrypt_batch.py "$env:USERPROFILE\Desktop\CONVERSION IMAGE" "$env:USERPROFILE\Downloads\decode"
```

Use an output folder on a drive that supports hard links (such as NTFS). The batch command verifies and writes original JPEG bytes, preserves existing output files and reports failed inputs. For a different client batch, set its authorized key before launching the command.

Scroll the loaded scans list and click a scan to open it. Use the preview scrollbars to see the whole page. Export Excel processes every loaded scan across all selected folders. Scan review lists every input image, including any with read errors or zero matches.

## Local source workbook for Excel export

Keep the original KYC MI workbook as `KYC_MI_Source.xlsx` beside `app.py` or beside the packaged Workbench executable for automatic discovery. A valid remembered source or a valid workbook imported in the current session takes priority. Use **File > Change source workbook...** to select a different source. Output is saved in **Downloads > KYC Exports**. Keep the workbook local because it contains KYC records.
