package rpa_relatorio.rpa_relatorio.Service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import rpa_relatorio.rpa_relatorio.Repository.RelatorioRepository09h;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

@Service
public class RoboService {

    @Autowired
    public RelatorioRepository09h repository;

    @Value("${report.output.directory}")
    public String outputDirectory;

    private static final Logger logger = LoggerFactory.getLogger(RoboService.class);

    @Scheduled(cron = "0 39 23 * * *")
    public void agendamentoSisfies09h() {
        processarRelatorio("SISFIES", "09h");
    }

    @Scheduled(cron = "0 27 23 * * *")
    public void agendamentoSisfies15h() {
        processarRelatorio("SISFIES", "15h");
    }

    @Scheduled(cron = "0 27 23 * * *")
    public void agendamentoSisfies17h() {
        processarRelatorio("SISFIES", "17h");
    }

    @Scheduled(cron = "0 39 23 * * *")
    public void agendamentoSisprouni09h() {
        processarRelatorio("SISPROUNI", "09h");
    }

    @Scheduled(cron = "0 28 23 * * *")
    public void agendamentoSisprouni15h() {
        processarRelatorio("SISPROUNI", "15h");
    }

    @Scheduled(cron = "0 28 23 * * *")
    public void agendamentoSisprouni17h() {
        processarRelatorio("SISPROUNI", "17h");
    }

    public void processarRelatorio(String tipoRelatorio, String horaExecucao) {
        logger.info("⏳ Iniciando processamento do relatório [{}] às {}", tipoRelatorio, horaExecucao);
        try {
            List<Object[]> resultados = gerarRelatorio(tipoRelatorio);
            String nomeArquivo = nomearRelatorio(tipoRelatorio, horaExecucao);

            logger.info("📄 Gerando arquivo TXT: {}", nomeArquivo);
            salvarArquivoTXT(resultados, nomeArquivo);

            logger.info("📊 Gerando arquivo Excel");
            exportarParaExcel(nomeArquivo, tipoRelatorio, horaExecucao);

            logger.info("✅ Relatório [{}] às {} finalizado com sucesso!", tipoRelatorio, horaExecucao);

        } catch (Exception e) {
            logger.error("❌ Erro ao gerar relatório {} às {}: {}", tipoRelatorio, horaExecucao, e.getMessage(), e);
        }
    }

    private List<Object[]> gerarRelatorio(String tipoRelatorio) throws Exception {
        if ("SISFIES".equalsIgnoreCase(tipoRelatorio)) {
            return repository.SISFIES();
        } else if ("SISPROUNI".equalsIgnoreCase(tipoRelatorio)) {
            return repository.SISPROUNI();
        } else {
            throw new IllegalArgumentException("Tipo de relatório desconhecido: " + tipoRelatorio);
        }
    }

    private String nomearRelatorio(String tipoRelatorio, String horaExecucao) {
        String data = LocalDate.now().format(DateTimeFormatter.ofPattern("ddMMyyyy"));
        String nomeArquivo = "Relatorio-Documentos-" + tipoRelatorio.toUpperCase() + data + "(bases-24.2)-" + horaExecucao + ".txt";

        File pasta = new File(outputDirectory);
        if (!pasta.exists()) {
            pasta.mkdirs();
        }

        return outputDirectory + File.separator + nomeArquivo;
    }

    private void salvarArquivoTXT(List<Object[]> resultados, String nomeArquivo) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(nomeArquivo), StandardCharsets.UTF_8))) {
            for (Object[] linha : resultados) {
                StringBuilder linhaTexto = new StringBuilder();
                for (Object valor : linha) {
                    linhaTexto.append(valor != null ? valor.toString() : "").append(";");
                }
                writer.write(linhaTexto.toString().replaceAll(";$", "")); // remove último ;
                writer.newLine();
            }
        }
    }

    public void exportarParaExcel(String nomeArquivoTXT, String tipoRelatorio, String horaExecucao) throws Exception {
        String data = LocalDate.now().format(DateTimeFormatter.ofPattern("ddMMyyyy"));
        String nomeArquivoXlsx = outputDirectory + File.separator +
                "Relatorio-Documentos-" + tipoRelatorio.toUpperCase() + data + "(bases-24.2)-" + horaExecucao + ".xlsx";

        try (
                BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(nomeArquivoTXT), StandardCharsets.UTF_8));
                Workbook workbook = new XSSFWorkbook()
        ) {
            Sheet sheet = workbook.createSheet(tipoRelatorio.toUpperCase());

            List<String> cabecalho = getCabecalho(tipoRelatorio);
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < cabecalho.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(cabecalho.get(i));

                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                style.setFont(font);
                cell.setCellStyle(style);
            }

            String linha;
            int linhaIndex = 1;
            while ((linha = reader.readLine()) != null) {
                String[] valores = linha.split(";");
                Row row = sheet.createRow(linhaIndex++);
                for (int i = 0; i < valores.length; i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(valores[i].trim());
                }
            }

            for (int i = 0; i < cabecalho.size(); i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream fos = new FileOutputStream(nomeArquivoXlsx)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            logger.error("Erro ao exportar arquivo Excel: {}", e.getMessage(), e);
            throw e;
        }
    }
    public List<String> getCabecalho(String tipo) {
        List<String> cabecalho = new ArrayList<>();
        if ("SISPROUNI".equalsIgnoreCase(tipo)) {
            cabecalho.add("ID");
            cabecalho.add("Data Criação");
            cabecalho.add("Documento Id");
            cabecalho.add("Situação");
            cabecalho.add("Regional");
            cabecalho.add("Cod. Instituição");
            cabecalho.add("Instituição");
            cabecalho.add("Cod. Campus");
            cabecalho.add("Campus");
            cabecalho.add("Cod. Curso");
            cabecalho.add("Curso");
            cabecalho.add("Aluno");
            cabecalho.add("CPF");
            cabecalho.add("Nome Importação");
            cabecalho.add("Numero Candidato");
            cabecalho.add("Numero Inscrição");
            cabecalho.add("Matrícula");
            cabecalho.add("Local Oferta");
            cabecalho.add("Chamada");
            cabecalho.add("Forma Ingresso");
            cabecalho.add("Periodo Ingresso");
            cabecalho.add("Tipo de Processo");
            cabecalho.add("Documento");
            cabecalho.add("Número Membro");
            cabecalho.add("Status");
            cabecalho.add("Irregularidade");
            cabecalho.add("Observação");
            cabecalho.add("Analista");
            cabecalho.add("Analista Login");
            cabecalho.add("Situação Anterior");
            cabecalho.add("Data Envio Análise");
            cabecalho.add("Data Finalização Análise");
            cabecalho.add("Número de Página");
            cabecalho.add("Professor de Rede Pública");
            cabecalho.add("Ensino Médio Em");
            cabecalho.add("Candidato e Deficiente");
            cabecalho.add("Polo Parceiro");
            cabecalho.add("Tipo de Bolsa Importação");
            cabecalho.add("Turno Importação");
            cabecalho.add("Endereço Importação");
            cabecalho.add("Cidade Importação");
            cabecalho.add("Cep Importação");
            cabecalho.add("E-mail Importação");
            cabecalho.add("DDD Telefone Importação");
            cabecalho.add("Nota Média Importação");
            cabecalho.add("Tipo Prouni Importação");
            cabecalho.add("Curso Importação");
            cabecalho.add("CPF Importação");
            cabecalho.add("Período Importação");
            cabecalho.add("Data Vínculo");
            cabecalho.add("Pasta Vermelha");
            cabecalho.add("Usa Termo");
            cabecalho.add("Possui Formação Complementar ao Curso Selecionado?");
            cabecalho.add("Curso de Formação");
            cabecalho.add("Habilitação");
            cabecalho.add("Multiplicador");
            cabecalho.add("Limite Salário Familiar");
            cabecalho.add("Renda per Capita");
            cabecalho.add("Qtde Salários Mínimos");
            cabecalho.add("Validador");
            cabecalho.add("Resultado SisProuni");
            cabecalho.add("Documentos Mínimos");
            cabecalho.add("Classificação");
            cabecalho.add("1 - Raça / Cor do Candidato");
            cabecalho.add("5 - Vínculo com Ies Pública?");
            cabecalho.add("6 - Formação Complementar ao Curso Classificado?");
            cabecalho.add("6.2 - Habilitação do Curso de Formação");
            cabecalho.add("6.1 - Curso de Formação");
            cabecalho.add("Obrigatoriedade Doc");
            cabecalho.add("Modalidade");
        } else if ("SISFIES".equalsIgnoreCase(tipo)) {
            cabecalho.addAll(getCabecalho("SISPROUNI")); // Reutiliza o cabeçalho do SISPROUNI
        }
        return cabecalho;
    }
}
