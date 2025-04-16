package rpa_relatorio.rpa_relatorio.Repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;
import rpa_relatorio.rpa_relatorio.Entity.Relatorio13h;


import java.util.List;

@Repository
public interface RelatorioRepository13h extends JpaRepository<Relatorio13h, Long> {

    @Query(value = "SELECT * FROM processo", nativeQuery = true)
    List<Object[]> APROVADOS ();


    @Query(value = "SELECT * FROM processo", nativeQuery = true)
    List<Object[]> PENDENCIADOS ();
}
