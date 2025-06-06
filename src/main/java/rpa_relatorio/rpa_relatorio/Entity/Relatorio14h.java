package rpa_relatorio.rpa_relatorio.Entity;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import lombok.Data;

import java.time.LocalDateTime;

@Entity
@Data
public class Relatorio14h {

    @Id
    private Long id;

    private LocalDateTime dataGeracao;
    private String tipoRelatorio;
    private String horaExecucao;
}
