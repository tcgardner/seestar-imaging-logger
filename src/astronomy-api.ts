import https from 'node:https';
import { ASTRONOMICAL_OBJECTS } from './astronomical-objects.js';

/**
 * Cache for object type lookups to avoid repeated API calls
 */
let objectTypeCache = new Map<string, string>();

/**
 * Comprehensive SIMBAD object type code mapping
 * Maps astronomical object codes to categories for Google Sheets color-coding
 */
const SIMBAD_OTYPE_MAP: Record<string, { category: string; description: string }> = {
  // Galaxies
  'G': { category: 'Galaxy', description: 'Galaxy' },
  'GG': { category: 'Galaxy', description: 'Group of galaxies' },
  'GiG': { category: 'Galaxy', description: 'Giant elliptical galaxy' },
  'GiP': { category: 'Galaxy', description: 'Giant spiral galaxy' },
  'Sy': { category: 'Galaxy', description: 'Seyfert' },
  'Sy1': { category: 'Galaxy', description: 'Seyfert 1' },
  'Sy2': { category: 'Galaxy', description: 'Seyfert 2' },
  'QSO': { category: 'Galaxy', description: 'Quasar' },
  'AGN': { category: 'Galaxy', description: 'Active Galactic Nucleus' },
  'LINER': { category: 'Galaxy', description: 'LINER galaxy' },
  'LIN': { category: 'Galaxy', description: 'LINER' },
  'Sy?': { category: 'Galaxy', description: 'Seyfert?' },
  'Bla': { category: 'Galaxy', description: 'Blazar' },
  'BLL': { category: 'Galaxy', description: 'BL Lac object' },
  'OVV': { category: 'Galaxy', description: 'Optically Violent Variable' },
  'grav': { category: 'Galaxy', description: 'Gravitational lens' },

  // Nebulae
  'PN': { category: 'Nebula', description: 'Planetary nebula' },
  'SNR': { category: 'Nebula', description: 'Supernova remnant' },
  'HII': { category: 'Nebula', description: 'HII region' },
  'EmO': { category: 'Nebula', description: 'Emission object' },
  'Neb': { category: 'Nebula', description: 'Nebula' },
  'RNe': { category: 'Nebula', description: 'Reflection nebula' },
  'DNe': { category: 'Nebula', description: 'Dark nebula' },
  'MoC': { category: 'Nebula', description: 'Molecular cloud' },
  'CGG': { category: 'Nebula', description: 'Compact group of galaxies' },
  'C*': { category: 'Nebula', description: 'Cluster of stars' },
  'H2O': { category: 'Nebula', description: 'H2O maser' },
  'OH': { category: 'Nebula', description: 'OH maser' },
  'CH3OH': { category: 'Nebula', description: 'CH3OH maser' },
  'SFR': { category: 'Nebula', description: 'Star forming region' },
  'Cld': { category: 'Nebula', description: 'Cloud' },
  'GNe': { category: 'Nebula', description: 'Gaseous nebula' },
  'BN': { category: 'Nebula', description: 'Brown dwarf nebula' },

  // Star Clusters
  'Cl*': { category: 'Star Cluster', description: 'Star cluster' },
  'GlC': { category: 'Star Cluster', description: 'Globular cluster' },
  'OpC': { category: 'Star Cluster', description: 'Open cluster' },
  'Ass': { category: 'Star Cluster', description: 'Association' },
  'St': { category: 'Star Cluster', description: 'Star' },
  'Cl': { category: 'Star Cluster', description: 'Cluster' },
  'Movi': { category: 'Star Cluster', description: 'Moving group' },

  // Solar System
  'Planet': { category: 'Planet', description: 'Planet' },
  'Moon': { category: 'Moon', description: 'Moon' },
  'Sun': { category: 'Star', description: 'Sun' },
  'Comet': { category: 'Other', description: 'Comet' },
  'Asteroid': { category: 'Other', description: 'Asteroid' },

  // Stars
  '*': { category: 'Star', description: 'Star' },
  'V*': { category: 'Star', description: 'Variable star' },
  'Puls': { category: 'Star', description: 'Pulsating variable' },
  'RR*': { category: 'Star', description: 'RR Lyrae variable' },
  'Mira': { category: 'Star', description: 'Mira variable' },
  'CV*': { category: 'Star', description: 'Cataclysmic variable' },
  'WD*': { category: 'Star', description: 'White dwarf' },
  'N*': { category: 'Star', description: 'Neutron star' },
  'BH': { category: 'Star', description: 'Black hole' },
  'BH?': { category: 'Star', description: 'Black hole candidate' },
  'WR*': { category: 'Star', description: 'Wolf-Rayet star' },
  'Be*': { category: 'Star', description: 'Be star' },
  'XB*': { category: 'Star', description: 'X-ray binary' },
  'LXB': { category: 'Star', description: 'Low-mass X-ray binary' },
  'HXB': { category: 'Star', description: 'High-mass X-ray binary' },
  'Pulsar': { category: 'Star', description: 'Pulsar' },
  'AGB*': { category: 'Star', description: 'AGB star' },
  'OH/IR': { category: 'Star', description: 'OH/IR star' },
  'Symbi': { category: 'Star', description: 'Symbiotic star' },

  // Misc
  'MWG': { category: 'Other', description: 'Milky Way Galaxy' },
  'MW': { category: 'Other', description: 'Milky Way' },
  'ISM': { category: 'Nebula', description: 'Interstellar medium' },
  'rG': { category: 'Galaxy', description: 'Radio galaxy' },
  'EmG': { category: 'Galaxy', description: 'Emission line galaxy' },
  'Narrow': { category: 'Nebula', description: 'Narrow-line region' },
  'mv': { category: 'Galaxy', description: 'Molecular absorption' },
};

/**
 * Uses the official CDS Sesame Name Resolver (aggregates SIMBAD + NED + VizieR)
 * to determine object type from catalog identifier (M31, NGC7000, etc.).
 * Returns one of the categories used by the Google Sheets color rules.
 */
export async function getObjectTypeFromCatalog(catalog: string): Promise<string> {
  if (!catalog?.trim()) {
    return 'Other---';
  }

  const key = catalog.trim().toUpperCase();
  if (objectTypeCache.has(key)) {
    return objectTypeCache.get(key)!;
  }

  // Sesame machine-readable endpoint (plain-text output)
  const url = `https://cds.unistra.fr/cgi-bin/nph-sesame?${encodeURIComponent(key)}`;

  return new Promise((resolve) => {
    https.get(url, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode !== 200) {
          objectTypeCache.set(key, 'Other');
          resolve('Other');
          return;
        }

        // Sesame returns lines like: %C.0 AGN   or   %C.0 HII   or   %C.0 Cl*
        const lines = data.split('\n');
        let otypeCode = '';

        for (const line of lines) {
          if (line.includes('%C.0 ')) {
            const after = line.split('%C.0 ')[1];
            if (after) {
              otypeCode = after.trim().split(/\s+/)[0];
              break;
            }
          }
        }

        // Map Sesame otype codes using comprehensive lookup table
        let mapped = `Other`;
        if (otypeCode) {
          const mapping = SIMBAD_OTYPE_MAP[otypeCode];
          if (mapping) {
            mapped = `${mapping.category} (${mapping.description})`;
          } else {
            // Fallback for unmapped codes
            mapped = `Other`;
          }
        }

        objectTypeCache.set(key, mapped);
        resolve(mapped);
      });
    }).on('error', (error) => {
      objectTypeCache.set(key, 'Other');
      resolve('Other');
    });
  });
}

/**
 * Formats catalog code to human-readable name.
 * Looks up in astronomical objects table, falls back to formatted code.
 * Examples: M42 → "Messier 42 - Orion Nebula", IC434 → "IC 434"
 * 
 * Note: The ASTRONOMICAL_OBJECTS lookup table is imported from astronomical-objects.ts
 */
export function getObjectNameFromCatalog(catalog: string): string {
  if (!catalog?.trim()) {
    return '';
  }

  const key = catalog.replace(/\s+/g, '').toUpperCase();
  
  // Check if exact match in lookup table
  if (ASTRONOMICAL_OBJECTS[key]) {
    return ASTRONOMICAL_OBJECTS[key].name;
  }

  // Parse and format the catalog code
  const match = key.match(/^([A-Z]+)\s*(\d+)$/);
  if (match) {
    const [, prefix, number] = match;
    
    // Format common prefixes
    const prefixMap: Record<string, string> = {
      'M': 'Messier',
      'NGC': 'NGC',
      'IC': 'IC',
      'UGC': 'UGC',
      'PGC': 'PGC',
      'ABELL': 'Abell',
      'CALDWELL': 'Caldwell',
    };
    
    const formattedPrefix = prefixMap[prefix] || prefix;
    return `${formattedPrefix} ${number}`;
  }

  // Return as-is if no match
  return catalog;
}

/**
 * Clear the object type cache (useful for testing or memory management)
 */
export function clearObjectTypeCache(): void {
  objectTypeCache.clear();
}

/**
 * Check if an astronomical object is visible from North America
 * Returns true if visible, false otherwise
 */
export function isVisibleInNorthAmerica(catalog: string): boolean {
  if (!catalog?.trim()) {
    return false;
  }

  const key = catalog.replace(/\s+/g, '').toUpperCase();
  const obj = ASTRONOMICAL_OBJECTS[key];
  
  return obj ? obj.visibleInNA : false;
}
