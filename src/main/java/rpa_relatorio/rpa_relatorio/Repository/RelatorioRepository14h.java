package rpa_relatorio.rpa_relatorio.Repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;
import rpa_relatorio.rpa_relatorio.Entity.Relatorio14h;


import java.util.List;

@Repository
public interface RelatorioRepository14h extends JpaRepository<Relatorio14h, Long> {

    @Query(value = "select * from Relatorio", nativeQuery = true)
    List<Object[]> SISPROUNI ();


    @Query(value = "select * from Relatorio", nativeQuery = true)
    List<Object[]> SISFIES ();
}
