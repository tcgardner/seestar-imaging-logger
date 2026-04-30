/**
 * FITS Header Parser - Complete version with all requested + previous fields
 * Supports primary HDU and extensions (headers only)
 */

export interface FitsHeader {
    [keyword: string]: string | number | boolean | null | undefined;
  }
  
  export interface FitsHDU {
    header: FitsHeader;
    headerSize: number;   // always multiple of 2880 bytes
    dataOffset: number;
  }
  
  export interface StandardFitsHeader {
    // Core mandatory keywords
    SIMPLE: boolean;
    BITPIX: number;
    NAXIS: number;
    NAXIS1: number;
    NAXIS2: number;
    NAXISn: number[];
  
    // Exposure & stacking (all common variants)
    EXPTIME?: number;     // single frame exposure time (seconds)
    EXPOSURE?: number;    // alias for EXPTIME
    TOTALEXP?: number;    // total integrated exposure after stacking
    stackCount: number;
  
    // Observation / target info
    OBJECT?: string;
    CATALOG?: string;
    FILTER?: string;
  
    // Camera / detector
    GAIN?: number;
  
    // Telescope & Instrument
    TELESCOP?: string;
    INSTRUME?: string;
  
    // World Coordinates (WCS)
    RA?: number;          // Right Ascension (degrees)
    DEC?: number;         // Declination (degrees)
    CRVAL1?: number;      // Usually RA
    CRVAL2?: number;      // Usually Dec
  
    // Dates
    DATE?: string;
    DATE_OBS?: string;
  
    // Data scaling & units
    BSCALE?: number;
    BZERO?: number;
    BUNIT?: string;
  
    // Other common fields kept from earlier versions
    EXTEND?: boolean;
    SNAPSHOT?: number;
    NCOMBINE?: number;
    NEXP?: number;
    NFRAMES?: number;
    STACKCNT?: number;
  }
  
  /**
   * Parse all HDUs headers from ArrayBuffer
   */
  export function parseFitsHeaders(buffer: ArrayBuffer): FitsHDU[] {
    const view = new DataView(buffer);
    const hdus: FitsHDU[] = [];
    let offset = 0;
  
    while (offset + 2880 <= buffer.byteLength) {
      const { header, headerBytes } = parseSingleHeader(view, offset);
      hdus.push({
        header,
        headerSize: headerBytes,
        dataOffset: offset + headerBytes,
      });
      offset += headerBytes;
    }
  
    return hdus;
  }
  
  function parseSingleHeader(view: DataView, startOffset: number): { header: FitsHeader; headerBytes: number } {
    const header: FitsHeader = {};
    let offset = startOffset;
    let cards = 0;
  
    while (true) {
      if (offset + 80 > view.byteLength) break;
  
      const card = readCard(view, offset);
      offset += 80;
      cards++;
  
      if (card.keyword === "END") break;
  
      if (card.keyword && card.value !== undefined) {
        if (!(card.keyword in header)) {
          header[card.keyword] = card.value;
        }
      }
  
      if (cards > 10000) break; // prevent malformed files
    }
  
    const headerBytes = Math.ceil((offset - startOffset) / 2880) * 2880;
    return { header, headerBytes };
  }
  
  function readCard(view: DataView, offset: number) {
    let cardStr = "";
    for (let i = 0; i < 80; i++) {
      cardStr += String.fromCharCode(view.getUint8(offset + i));
    }
  
    const keyword = cardStr.substring(0, 8).trim();
  
    if (keyword === "END") {
      return { keyword, value: null };
    }
  
    if (cardStr.substring(8, 10) !== "= ") {
      return { keyword, value: null };
    }
  
    let valueField = cardStr.substring(10).trim();
    const slashPos = valueField.indexOf("/");
    if (slashPos !== -1) {
      valueField = valueField.substring(0, slashPos).trim();
    }
  
    let value: any = null;
    if (valueField === "T") value = true;
    else if (valueField === "F") value = false;
    else if (valueField.startsWith("'") && valueField.endsWith("'")) {
      value = valueField.slice(1, -1).trim();
    } else if (/^-?\d+$/.test(valueField)) {
      value = parseInt(valueField, 10);
    } else if (/^-?[\d.eE+-]+$/.test(valueField)) {
      value = parseFloat(valueField);
    } else {
      value = valueField;
    }
  
    return { keyword, value };
  }
  
  /**
   * Extract all fields with safe fallbacks
   */
  export function getStandardHeader(hdu: FitsHDU): StandardFitsHeader {
    const h = hdu.header;
  
    const naxis = Number(h.NAXIS ?? 0);
    const naxisn: number[] = [];
    for (let i = 1; i <= naxis; i++) {
      naxisn.push(Number(h[`NAXIS${i}`] ?? 0));
    }
  
    // Stack / combine count
    const stackCount = Number(
      h.NCOMBINE ?? h.STACKCNT ?? h.SNAPSHOT ?? h.NEXP ?? h.NFRAMES ?? 1
    );
  
    // Exposure times (single frame)
    const exptime = Number(h.EXPTIME ?? h.EXPOSURE ?? 0);
  
    // Total exposure (try direct total first, then compute from single × stack)
    const totalexp = Number(
      h.TOTALEXP ?? h["TOTAL-EXP"] ?? h["TOTEXP"] ?? (exptime * stackCount)
    );
  
    // Coordinates
    const ra = Number(h.RA ?? h.CRVAL1 ?? 0);
    const dec = Number(h.DEC ?? h.CRVAL2 ?? 0);
  
    return {
      SIMPLE: Boolean(h.SIMPLE ?? false),
      BITPIX: Number(h.BITPIX ?? 0),
      NAXIS: naxis,
      NAXIS1: naxisn[0] ?? 0,
      NAXIS2: naxisn[1] ?? 0,
      NAXISn: naxisn,
  
      // Exposure
      EXPTIME: exptime > 0 ? exptime : undefined,
      EXPOSURE: exptime > 0 ? exptime : undefined,
      TOTALEXP: totalexp > 0 ? totalexp : undefined,
      stackCount,
  
      // Observation
      OBJECT: String(h.OBJECT ?? h.OBJ ?? ""),
      CATALOG: String(h.CATALOG ?? h.CAT ?? ""),
      FILTER: String(h.FILTER ?? h.FILT ?? ""),
  
      // Camera
      GAIN: h.GAIN != null ? Number(h.GAIN) : undefined,
  
      // Telescope / Instrument
      TELESCOP: String(h.TELESCOP ?? ""),
      INSTRUME: String(h.INSTRUME ?? ""),
  
      // Coordinates
      RA: ra !== 0 ? ra : undefined,
      DEC: dec !== 0 ? dec : undefined,
      CRVAL1: h.CRVAL1 != null ? Number(h.CRVAL1) : undefined,
      CRVAL2: h.CRVAL2 != null ? Number(h.CRVAL2) : undefined,
  
      // Dates & scaling
      DATE: String(h.DATE ?? ""),
      DATE_OBS: String(h["DATE-OBS"] ?? h.DATE_OBS ?? ""),
      BSCALE: h.BSCALE != null ? Number(h.BSCALE) : undefined,
      BZERO: h.BZERO != null ? Number(h.BZERO) : undefined,
      BUNIT: String(h.BUNIT ?? ""),
  
      // Keep previous common fields for compatibility
      EXTEND: h.EXTEND as boolean | undefined,
      SNAPSHOT: h.SNAPSHOT as number | undefined,
      NCOMBINE: h.NCOMBINE as number | undefined,
      NEXP: h.NEXP as number | undefined,
      NFRAMES: h.NFRAMES as number | undefined,
      STACKCNT: h.STACKCNT as number | undefined,
    };
  }
  
  // Convenience function
  export async function readFitsHeadersFromFile(file: File | Blob): Promise<FitsHDU[]> {
    const buffer = await file.arrayBuffer();
    return parseFitsHeaders(buffer);
  }