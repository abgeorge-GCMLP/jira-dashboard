const fs = require('fs');
const raw = fs.readFileSync('C:/Users/ABGEORGE/pcap_tsv2.txt', 'utf8');

const CANOE_BASE = 'https://client.canoesoftware.com/frontend/dataintelligence/document-center/document/';

const lines = raw.split('\n')
  .map(l => l.replace(/\r/g, '').trim())
  .filter(l => l && !l.match(/^[-|]+$/) && !l.startsWith('('));

const dataLines = lines.slice(1);

const mapped = dataLines.map(l => {
  const parts = l.split('|');
  if (parts.length < 11) return null;
  const q4id   = parts[4] !== 'NULL' ? parts[4].trim() : null;
  const q4dt   = parts[5] !== 'NULL' ? parts[5].trim() : null;
  const q4tm   = parts[6] !== 'NULL' ? parts[6].trim() : null;
  const q1id   = parts[7] !== 'NULL' ? parts[7].trim() : null;
  const q1dt   = parts[8] !== 'NULL' ? parts[8].trim() : null;
  const q1tm   = parts[9] !== 'NULL' ? parts[9].trim() : null;
  return {
    deal:   parts[0],
    leId:   parseInt(parts[1]),
    le:     parts[2],
    docType:parts[3],
    q4id,   q4url: q4id ? CANOE_BASE + q4id : null,
    q4date: q4dt,  q4time: q4tm,
    q1id,   q1url: q1id ? CANOE_BASE + q1id : null,
    q1date: q1dt,  q1time: q1tm,
    q1: parts[10] === 'R' ? 'Received' : 'Not Received'
  };
}).filter(r => r && r.deal && r.le && r.docType && !isNaN(r.leId));

fs.writeFileSync('C:/Users/ABGEORGE/pcap_data.json', JSON.stringify(mapped));
console.log('Total:', mapped.length,
  '| Q4 times:', mapped.filter(r=>r.q4time).length,
  '| Q1 times:', mapped.filter(r=>r.q1time).length);
