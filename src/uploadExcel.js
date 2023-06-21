import React, { useState } from "react";
import axios from "axios";
import { read, utils, write, writeFile } from "xlsx";
import gif from "./gifs-de-aguarde-0.gif";

const App = () => {
  const [file, setFile] = useState(null);
  const [resultFile, setResultFile] = useState(null);
  const [carregando, setCarregando] = useState(false);
  const handleFileUpload = (event) => {
    const uploadedFile = event.target.files[0];
    setFile(uploadedFile);
  };

  const calculateDistance = async () => {
    if (file === null) {
      return;
    }
    setCarregando(true);
    try {
      const fileReader = new FileReader();
      fileReader.onload = async (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = read(data, { type: "array" });

        // Assume que a planilha está na primeira folha (Sheet1)
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = utils.sheet_to_json(worksheet, { header: 1 });

        // Itera sobre as linhas da planilha
        const updatedData = await Promise.all(
          jsonData.map(async (row, rowIndex) => {
            if (rowIndex === 0) {
              // Mantém o cabeçalho intacto e adiciona a coluna "distancia"
              row.push("distancia");
              return row;
            }

            const cepLoja = row[22]; // Coluna do CEP da Loja
            const cepCliente = row[12]; // Coluna do CEP do Cliente

            // Chama a API do servidor Node.js para obter a distância
            const url = `https://calcula.onrender.com/api/maps/api/distancematrix/json?origins=${cepLoja}&destinations=${cepCliente}`;
            try {
              const response = await axios.get(url);
              const distance = response.data.rows[0].elements[0].distance.text;
              row.push(distance);
            } catch (error) {
              row.push("Erro ao calcular a distância");
            }

            return row;
          })
        );

        // Cria um novo workbook com os dados atualizados
        const updatedWorkbook = utils.book_new();
        const updatedWorksheet = utils.aoa_to_sheet(updatedData);
        utils.book_append_sheet(updatedWorkbook, updatedWorksheet, "Sheet1");

        // Gera um blob a partir do workbook
        const arrayBuffer = write(updatedWorkbook, {
          type: "array",
          bookType: "xlsx",
        });

        // Cria um novo arquivo resultante com a coluna de distância adicionada
        const resultWorkbook = new Blob([arrayBuffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        setCarregando(false);
        // Define o arquivo resultante
        setResultFile(resultWorkbook);
      };

      fileReader.readAsArrayBuffer(file);
    } catch (error) {
      console.log(error);
    }
  };

  const handleDownload = () => {
    const downloadUrl = URL.createObjectURL(resultFile);
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = "planilha_resultante.xlsx";
    link.click();
  };

  return (
    <>
      <h1 style={{ marginTop: "30vh" }}>Calculo de distância</h1>
      <div className="conteudo">
        <input
          className="btn btn-light"
          style={{ marginRight: "15px" }}
          type="file"
          accept=".xlsx"
          onChange={handleFileUpload}
        />
        <button className="btn btn-primary" onClick={calculateDistance}>
          Calcular Distância
        </button>
        {resultFile && (
          <button
            style={{ marginLeft: "15px" }}
            className="btn btn-success"
            onClick={handleDownload}
          >
            Download
          </button>
        )}
      </div>
      {carregando ? <img alt="carregando" src={gif} /> : <></>}
    </>
  );
};

export default App;
