# KYC Workbench

Run locally with Python 3.10 or newer and Tk. On Windows, the standard Python installer includes Tk by default. On Linux, install your distribution's Python Tk package if Tk is missing.

```shell
python -m pip install -r requirements.txt
python app.py
```

Choose **Open image** for one scan or **Open folder** for a batch. To combine batches, use **Open multiple folders**, click **Add folder** for each folder, and select **Open folders**. The app uses each folder's configured key and combines all rows into one Excel export. JPEGs saved from multiple folders go into numbered subfolders so repeated filenames stay separate. The supplied Client 2 key is built in. **Decrypt Images** saves a verified JPEG or saves verified JPEGs into a folder you choose. Use an output drive with hard link support, such as NTFS.

## Export records from Client 2 grid pages

1. Open your encrypted JPG folder and check that the page preview is readable.
2. Click **Export Excel**. If the source workbook picker opens, select the original numbered **KYC MI** workbook. Cancelling the picker cancels the export. This step uses local Tesseract OCR or the installed Windows OCR language.
3. The app reads printed record IDs and uses the matching MI rows to populate the 22 columns, with one row per matched ID on each image. Image Name holds the page number. The **Scan review** sheet lists the source pages, counts and unmatched IDs. Compare against the source pages because OCR may miss or misread IDs.

Use **File > Export OCR review...** to save raw OCR text and detected record IDs for troubleshooting. That workbook has **Scan review** and **Detected IDs** sheets.

For manually entered or labeled forms, use **File > Export reviewed form rows**. **Read scan text** can prefill explicitly labeled values, and **Check row** previews remarks.

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

Keep the original KYC MI workbook beside `app.py` as `KYC_MI_Source.xlsx` for automatic discovery. A valid remembered source or a valid workbook imported in the current session takes priority. Use **File > Change source workbook...** to select a different source. Output is saved in **Downloads > KYC Exports**. Keep the workbook local because it contains KYC records.
