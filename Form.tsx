import React from 'react';
import { useState } from 'react';
import { DocumentCreator } from './surat-generator';
import { saveAs } from 'file-saver';
import { Packer } from 'docx';

export default function Form() {
  const [patientName, setPatientName] = useState('');
  const [berat, setBerat] = useState('50');
  const [tinggi, setTinggi] = useState('150');
  const beratAsNumber = Number(berat);
  const tinggiAsNumber = Number(tinggi);

  function generate(nomor, nama, dokter, berat, tinggi): void {
    const documentCreator = new DocumentCreator();
    const doc = documentCreator.create([nomor, nama, dokter, berat, tinggi]);

    Packer.toBlob(doc).then((blob) => {
      console.log(blob);
      saveAs(blob, `${nomor}.${nama}.docx`);
      console.log('Document created successfully');
    });
  }

  return (
    <div>
      <label>
        First name:
        <input
          value={patientName}
          onChange={(e) => setPatientName(e.target.value)}
        />
      </label>
      <p></p>
      <label>
        Berat:
        <input
          value={berat}
          onChange={(e) => setBerat(e.target.value)}
          type="number"
        />
      </label>
      <p></p>
      <label>
        Tinggi:
        <input
          value={tinggi}
          onChange={(e) => setTinggi(e.target.value)}
          type="number"
        />
      </label>
      {patientName !== '' && <p>Namamu adalah {patientName}.</p>}
      {beratAsNumber > 0 && <p>Beratmu {beratAsNumber}kg.</p>}
      {tinggiAsNumber > 0 && <p>Tinggimu {tinggiAsNumber}cm.</p>}

      <p>
        <button
          onClick={() =>
            generate(1, patientName, 'lili', beratAsNumber, tinggiAsNumber)
          }
        >
          Generate CV with docx!
        </button>
      </p>
    </div>
  );
}
