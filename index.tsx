import React, { Component } from 'react';
import { useState } from 'react';
import { render } from 'react-dom';
import Hello from './Hello';
import Form from './Form';
import './style.css';

import { saveAs } from 'file-saver';
import { Packer } from 'docx';
import { experiences, education, skills, achievements } from './cv-data';
//import { nomor, nama, dokter, berat, tinggi } from './surat-data';
import { DocumentCreator } from './surat-generator';

interface AppProps {}
interface AppState {
  nama: string;
  nomor: number;
}

class App extends Component<AppProps, AppState> {
  constructor(props) {
    super(props);
    this.state = {
      nama: 'React',
      nomor: 10,
    };
  }

  render() {
    return (
      <div>
        <Hello name={this.state.nama} />
        <Form />
        <p>{__dirname}</p>
      </div>
    );
  }
}

render(<App />, document.getElementById('root'));
