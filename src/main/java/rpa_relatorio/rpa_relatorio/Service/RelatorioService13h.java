package rpa_relatorio.rpa_relatorio.Service;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import rpa_relatorio.rpa_relatorio.Config.NotificacaoTeams;
import rpa_relatorio.rpa_relatorio.Repository.RelatorioRepository13h;

import java.io.BufferedWriter;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Service
public class RelatorioService13h {

    @Autowired
    private NotificacaoTeams notificacaoTeams;

    @Autowired
    private RelatorioRepository13h repository;

    @Value("${report.output.directory}")
    public String outputDirectory;

    private static final Logger logger = LoggerFactory.getLogger(RelatorioService13h.class);

    @Scheduled(cron = "0 50 00 * * *")
    public void agendamentoAprovados13h() {
        processarRelatorio("APROVADOS", "13h");
    }

    @Scheduled(cron = "0 50 00 * * *")
    public void agendamentoPendenciados13h() {
        processarRelatorio("PENDENCIADOS", "13h");
    }

    public void processarRelatorio(String tipoRelatorio, String horaExecucao) {
        logger.info("Iniciando processamento do relatório [{}] às {}", tipoRelatorio, horaExecucao);

        try {
            List<Object[]> resultados = gerarRelatorio(tipoRelatorio);
            String nomeArquivo = nomearRelatorio(tipoRelatorio, horaExecucao);
            List<String> cabecalho = getCabecalho(tipoRelatorio);


            List<Map<String, Object>> dadosConvertidos = converterParaMap(resultados, cabecalho);

            logger.info("Total de registros antes do filtro: {}", dadosConvertidos.size());


            dadosConvertidos.forEach(linha -> logger.info("Linha antes do filtro: {}", linha));


            dadosConvertidos.removeIf(linha -> {
                Object tipoProcesso = linha.get(cabecalho.get(1));
                return tipoProcesso != null && tipoProcesso.toString().equalsIgnoreCase("Isenção de Disciplina");
            });

            logger.info("Total de registros após o filtro: {}", dadosConvertidos.size());

            dadosConvertidos.forEach(linha -> logger.info("Linha após o filtro: {}", linha));

            logger.info("Gerando arquivo CSV: {}", nomeArquivo);
            salvarArquivoCsv(dadosConvertidos, tipoRelatorio, nomeArquivo);

            logger.info("Relatório [{}] às {} finalizado com sucesso!", tipoRelatorio, horaExecucao);
            notificacaoTeams.enviarNotificacao("Relatório " +tipoRelatorio + " gerado com sucesso às " + horaExecucao + ".");

        } catch (Exception e) {
            logger.error("Falha ao gerar o relatório {} às {}: {}", tipoRelatorio, horaExecucao, e.getMessage(), e);
            notificacaoTeams.enviarNotificacao("Falha ao gerar o relatório " + tipoRelatorio + " às " + horaExecucao + ": " + e.getMessage());
        }
    }

    private List<Object[]> gerarRelatorio(String tipoRelatorio) throws Exception {
        if ("APROVADOS".equalsIgnoreCase(tipoRelatorio)) {
            return repository.APROVADOS();
        } else if ("PENDENCIADOS".equalsIgnoreCase(tipoRelatorio)) {
            return repository.PENDENCIADOS();
        } else {
            throw new IllegalArgumentException("Tipo de relatório desconhecido: " + tipoRelatorio);
        }
    }

    private String nomearRelatorio(String tipoRelatorio, String horaExecucao) {
        String data = LocalDate.now().format(DateTimeFormatter.ofPattern("ddMMyyyy"));
        String nomeArquivo = "Relatorio-Documentos-" + tipoRelatorio.toUpperCase() + data + "(bases-24.2)-13h.csv";

        File pasta = new File(outputDirectory);
        if (!pasta.exists()) {
            pasta.mkdirs();
        }

        return outputDirectory + File.separator + nomeArquivo;
    }

    public void salvarArquivoCsv(List<Map<String, Object>> dados, String tipo, String caminhoArquivo) {
        try (BufferedWriter writer = Files.newBufferedWriter(Paths.get(caminhoArquivo), StandardCharsets.UTF_8)) {
            if (!dados.isEmpty()) {
                List<String> cabecalho = getCabecalho(tipo);
                writer.write(String.join(";", cabecalho));
                writer.newLine();

                for (Map<String, Object> linha : dados) {
                    List<String> valores = new ArrayList<>();
                    for (String chave : cabecalho) {
                        String valor = String.valueOf(linha.getOrDefault(chave, ""));
                        valores.add(escapeCsvValue(valor));
                    }
                    writer.write(String.join(";", valores));
                    writer.newLine();
                }
                logger.info("CSV gerado com sucesso em: {}", caminhoArquivo);
            } else {
                logger.warn("Nenhum dado disponível para gerar o CSV.");
            }
        } catch (IOException e) {
            logger.error("Erro ao gerar o CSV para o tipo '{}' no caminho: {}", tipo, caminhoArquivo, e);
        }
    }

    private String escapeCsvValue(String valor) {
        if (valor.contains(";") || valor.contains("\"")) {
            valor = "\"" + valor.replace("\"", "\"\"") + "\"";
        }
        return valor;
    }

    private List<Map<String, Object>> converterParaMap(List<Object[]> resultados, List<String> cabecalho) {
        List<Map<String, Object>> listaMapeada = new ArrayList<>();

        for (Object[] linha : resultados) {
            Map<String, Object> mapa = new LinkedHashMap<>();
            for (int i = 0; i < cabecalho.size(); i++) {
                Object valor = i < linha.length ? linha[i] : null;
                mapa.put(cabecalho.get(i), valor);
            }
            listaMapeada.add(mapa);
        }

        return listaMapeada;
    }

    public List<String> getCabecalho(String tipo) {
        List<String> cabecalho = new ArrayList<>();
        if ("APROVADOS".equalsIgnoreCase(tipo)) {
            cabecalho.add("Processo_ID   ");
            cabecalho.add("Tipo_Processo ");
            cabecalho.add("Periodo_de_Ingresso  ");
            cabecalho.add("Numero_Candidato  ");
            cabecalho.add("Numero_Inscrição  ");
            cabecalho.add("CPF  ");
            cabecalho.add("Instituição  ");
            cabecalho.add("Cod. Campus  ");
            cabecalho.add("Campus  ");
            cabecalho.add("Nome_Candidato  ");
            cabecalho.add("Nome_Documento  ");
            cabecalho.add("Status_Documento  ");
            cabecalho.add("Origem  ");
            cabecalho.add("Data_Aprovação  ");
            cabecalho.add("Situação  ");
        } else if ("PENDENCIADOS".equalsIgnoreCase(tipo)) {
            cabecalho.add("Processo_ID  ");
            cabecalho.add("Tipo_Processo  ");
            cabecalho.add("Periodo_de_Ingresso  ");
            cabecalho.add("Numero_Candidato  ");
            cabecalho.add("Numero_Inscrição  ");
            cabecalho.add("CPF  ");
            cabecalho.add("Instituição  ");
            cabecalho.add("Cod. Campus  ");
            cabecalho.add("Campus  ");
            cabecalho.add("Data_Rejeite  ");
            cabecalho.add("Motivo_Rejeite  ");
            cabecalho.add("Status_Documento  ");
            cabecalho.add("Situação  ");
        }
        return cabecalho;
    }
    private List<String> ajustarCabecalhoVisual(List<String> cabecalhoOriginal) {
        List<String> ajustado = new ArrayList<>();
        for (String coluna : cabecalhoOriginal) {
            ajustado.add(coluna + "          "); // 5 espaços
        }
        return ajustado;
    }
}