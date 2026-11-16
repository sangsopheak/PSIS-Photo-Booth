# PSIS Photo Booth

A private, browser-based photo sheet maker for Paññāsāstra International School. Arrange student portraits into 28-photo A4 sheets at exactly 3.5 × 4.65 cm, then export to Word or PDF or print. Photos are processed locally in your browser and are not uploaded by the application.

## Open the website

**Live website:** https://sangsopheak.github.io/PSIS-Photo-Booth/

The first visit may take a minute after a deployment. GitHub builds and publishes the site automatically when changes are pushed to [main](https://github.com/sangsopheak/PSIS-Photo-Booth/tree/main).

## Run on your computer

1. Install Node.js 20 or newer.
2. Download or clone this repository.
3. Open a terminal in the project folder and run:

   ```sh
   npm install
   npm run dev
   ```

4. Open the local URL Vite prints, usually http://localhost:5173.

For a production build, run npm run build and npm run preview.

## Print workflow

1. Drag photos onto the upload area or choose multiple images.
2. Review the A4 preview and set copies, fit/fill, spacing, and cut lines.
3. Download Word or PDF, or use Print.
4. Print on A4 at **100% / Actual size**; disable “Fit to page” in the print dialog.

Supported images: JPG, JPEG, PNG, and WEBP. Processing and export happen on the device.
