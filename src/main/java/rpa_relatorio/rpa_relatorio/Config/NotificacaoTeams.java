package rpa_relatorio.rpa_relatorio.Config;

import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;
import org.springframework.http.ResponseEntity;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;

@Service
public class NotificacaoTeams {

    private final String webhookUrl = "";

    public void enviarNotificacao(String mensagem) {
        RestTemplate restTemplate = new RestTemplate();

        String payload = "{\"text\": \"" + mensagem + "\"}";

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_JSON);

        HttpEntity<String> request = new HttpEntity<>(payload, headers);
        ResponseEntity<String> response = restTemplate.postForEntity(webhookUrl, request, String.class);

        if (response.getStatusCode().is2xxSuccessful()) {
            System.out.println("Notificação enviada com sucesso!");
        } else {
            System.out.println("Falha ao enviar notificação: " + response.getStatusCode());
        }
    }
}
