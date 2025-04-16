package rpa_relatorio.rpa_relatorio.Service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import rpa_relatorio.rpa_relatorio.Config.NotificacaoTeams;
import rpa_relatorio.rpa_relatorio.Repository.RelatorioRepository14h;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

@Service
public class RelatorioService14h {

    @Autowired
    private NotificacaoTeams notificacaoTeams;

    @Autowired
    public RelatorioRepository14h repository;

    @Value("${report.output.directory}")
    private String outputDirectory;

    private static final Logger logger = LoggerFactory.getLogger(RelatorioService14h.class);

    @Scheduled(cron = "0 44 11 * * *")
    public void agendamentoSisfies14h() {
        processarRelatorio("SISFIES", "14h");
    }

    @Scheduled(cron = "0 50 00 * * *")
    public void agendamentoSisprouni14h() {
        processarRelatorio("SISPROUNI", "14h");
    }


    public void processarRelatorio(String tipoRelatorio, String horaExecucao) {
        logger.info("Iniciando processamento do relatório [{}] às {}", tipoRelatorio, horaExecucao);
        try {
            List<Object[]> resultados = gerarRelatorio(tipoRelatorio);
            String nomeArquivoXlsx = nomearRelatorio(tipoRelatorio, horaExecucao);

            exportarParaExcel(resultados, nomeArquivoXlsx, tipoRelatorio);

            logger.info("Relatório [{}] às {} finalizado com sucesso!", tipoRelatorio, horaExecucao);
            notificacaoTeams.enviarNotificacao("Relatório " + tipoRelatorio + " gerado com sucesso às " + horaExecucao + ".");

        } catch (Exception e) {
            logger.error("Falha ao gerar o relatório {} às {}: {}", tipoRelatorio, horaExecucao, e.getMessage(), e);
            notificacaoTeams.enviarNotificacao("Falha ao gerar o relatório " + tipoRelatorio + " às " + horaExecucao + ": " + e.getMessage());
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
        String nomeArquivo = "Relatorio-Documentos-" + tipoRelatorio.toUpperCase() + data + "(bases-24.1)-" + horaExecucao + ".xlsx";

        File pasta = new File(outputDirectory);
        if (!pasta.exists()) {
            pasta.mkdirs();
        }

        return outputDirectory + File.separator + nomeArquivo;
    }

    private void exportarParaExcel(List<Object[]> dados, String caminhoArquivo, String tipoRelatorio) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(tipoRelatorio.toUpperCase());

            List<String> cabecalho = getCabecalho(tipoRelatorio);

            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = workbook.createCellStyle();

            for (int i = 0; i < cabecalho.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(cabecalho.get(i));
                cell.setCellStyle(headerStyle);
            }

            CellStyle numberStyle = workbook.createCellStyle();
            DataFormat format = workbook.createDataFormat();
            numberStyle.setDataFormat(format.getFormat("0"));

            int linhaIndex = 1;
            for (Object[] linha : dados) {
                Row row = sheet.createRow(linhaIndex++);
                for (int i = 0; i < linha.length; i++) {
                    Cell cell = row.createCell(i);
                    String valor = linha[i] != null ? linha[i].toString().trim() : "";

                    if (i == 14 || i == 15 || i == 16) {
                        try {
                            long numero = Long.parseLong(valor);
                            cell.setCellValue(numero);
                            cell.setCellStyle(numberStyle);
                        } catch (Exception e) {
                            logger.warn("Erro ao processar número inteiro na linha {} coluna {}: {}", linhaIndex, i + 1, valor);
                            cell.setCellValue(valor);
                        }
                    } else {
                        cell.setCellValue(valor);
                    }
                }
            }
            for (int i = 0; i < cabecalho.size(); i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream fileOut = new FileOutputStream(caminhoArquivo)) {
                workbook.write(fileOut);
            }
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
            cabecalho.addAll(getCabecalho("SISPROUNI"));
        }
        return cabecalho;
    }
}
