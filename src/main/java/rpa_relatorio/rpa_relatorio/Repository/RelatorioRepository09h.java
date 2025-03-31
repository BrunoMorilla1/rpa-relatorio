package rpa_relatorio.rpa_relatorio.Repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;
import rpa_relatorio.rpa_relatorio.Entity.Relatorio09h;


import java.util.List;
import java.util.UUID;

@Repository
public interface RelatorioRepository09h extends JpaRepository<Relatorio09h, Long> {

    @Query(value = "select * from Relatorio", nativeQuery = true)
    List<Object[]> SISPROUNI ();


    @Query(value = "select * from Relatorio", nativeQuery = true)
    List<Object[]> SISFIES ();
}
