<?php
require './vendor/autoload.php';


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Configurações de conexão com o banco de dados
$servername = "localhost";
$username = "root";
$password = "root";
$dbname = "laravel";

// Cria uma conexão com o banco de dados
$conn = new mysqli($servername, $username, $password, $dbname);

// Verifica se a conexão foi estabelecida com sucesso
if ($conn->connect_error) {
    die("Erro na conexão com o banco de dados: " . $conn->connect_error);
}

// Consulta SQL para recuperar os dados da tabela "clientes"
$sql = "SELECT id, nome, endereco, cidade, cep, telefone FROM clientes";

// Executa a consulta
$result = $conn->query($sql);

// Crie uma nova instância da classe Spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Defina o cabeçalho da planilha
$sheet->setCellValue('A1', 'ID');
$sheet->setCellValue('B1', 'Nome');
$sheet->setCellValue('C1', 'Endereço');
$sheet->setCellValue('D1', 'Cidade');
$sheet->setCellValue('E1', 'CEP');
$sheet->setCellValue('F1', 'Telefone');

// Preencha a planilha com os dados do banco de dados
$row = 2;
while ($row_data = $result->fetch_assoc()) {
    $sheet->setCellValue('A' . $row, $row_data['id']);
    $sheet->setCellValue('B' . $row, $row_data['nome']);
    $sheet->setCellValue('C' . $row, $row_data['endereco']);
    $sheet->setCellValue('D' . $row, $row_data['cidade']);
    
    $cep = $row_data['cep'];
    $cep_formatado = substr($cep, 0, 5) . '-' . substr($cep, 5);
    $sheet->setCellValue('E' . $row, $cep_formatado);

    $telefone = $row_data['telefone'] ;
    $telefone_formatado = '(' . substr($telefone, 0, 2). ') '. substr($telefone, 2, 4). '-'. substr($telefone, 5) ;
    $sheet->setCellValue('F' . $row, $telefone_formatado);
    $row++;
}

// Salve a planilha em um arquivo Excel (xlsx)
$writer = new Xlsx($spreadsheet);
$filename = 'clientes.xlsx';
$writer->save($filename);

// Envie o arquivo Excel para download
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $filename . '"');
header('Cache-Control: max-age=0');

// Lembre-se de fechar a conexão com o banco de dados
$conn->close();
