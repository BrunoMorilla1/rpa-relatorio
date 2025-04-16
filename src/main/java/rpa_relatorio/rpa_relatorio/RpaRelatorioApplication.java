package rpa_relatorio.rpa_relatorio;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;

@SpringBootApplication
@EnableScheduling
public class RpaRelatorioApplication {

	public static void main(String[] args) {
		SpringApplication.run(RpaRelatorioApplication.class, args);
	}

}
