package br.com.cielo.newelo.grade.web.util;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import br.com.cielo.newelo.grade.interfaces.dto.LinhaGradeBrutaDTO;
import br.com.cielo.newelo.grade.web.bean.model.DadosGradeDTO;

/**
 * Classe para geração de arquivo XLS (Excel) a partir de objetos em lista
 *
 * @author CIT
 */
public class PlanilhaExcelUtil<T> {

    public byte[] geraPlanilha(final DadosGradeDTO dadosGradeDTO) throws IOException {
        byte[] retorno = null;

        List<LinhaGradeBrutaDTO> linhas = dadosGradeDTO.getGradeBruta().getLinhasDTO();

        if (linhas != null && linhas.size() > 0) {
            try (ByteArrayOutputStream saidaPlanilha = new ByteArrayOutputStream();
                            XSSFWorkbook workbook = new XSSFWorkbook();) {

                Sheet sheet = workbook.createSheet();

                int linha = 0;
                int coluna = 0;

                Row row = sheet.createRow(linha++);
                Cell cell;

                List<String> nomesColunas = Arrays.asList("Relatório Grade Bruta", "", "Data e Hora de geração",
                                "21/09/2018 15:26:33");

                for (String nomeColuna : nomesColunas) {
                    cell = row.createCell(coluna++);
                    cell.setCellValue(nomeColuna);
                }

                row = sheet.createRow(linha++);
                coluna = 0;

                nomesColunas = Arrays.asList("Bandeira", "Plataforma", "Data de Vencimento Inicial",
                                "Data de Vencimento Final", "Tipo de Pesquisa", "Emissor", "Credenciador");

                for (String nomeColuna : nomesColunas) {
                    cell = row.createCell(coluna++);
                    cell.setCellValue(nomeColuna);
                }

                row = sheet.createRow(linha++);
                coluna = 0;

                row.createCell(coluna++).setCellValue(7L);
                row.createCell(coluna++).setCellValue("D");
                row.createCell(coluna++).setCellValue(new Date());
                row.createCell(coluna++).setCellValue(new Date());
                row.createCell(coluna++).setCellValue("DETALHADO");
                row.createCell(coluna++).setCellValue("Todos");
                row.createCell(coluna++).setCellValue("Getnet");

                row = sheet.createRow(linha += 2);
                coluna = 0;

                this.geraPlanilhaLista(linha, sheet, (List<T>) linhas);

                /* gera os bytes do arquivo XLS */
                workbook.write(saidaPlanilha);
                retorno = saidaPlanilha.toByteArray();
            }

        }

        return retorno;
    }

    /**
     * Gera um arquivo de planilha Excel com os dados de uma lista de objetos
     *
     * @param nomeArquivo Nome do arquivo da planilha
     * @param linhasPlanilha Lista com
     * @throws IOException
     * @throws FileNotFoundException
     */
    private void geraPlanilhaLista(int linha, final Sheet sheet, final List<T> linhasPlanilha)
        throws IOException {

        List<String> nomesColunas = this.obtemNomeColunas();

        int coluna = 0;

        Row row = sheet.createRow(linha++);

        for (String nomeColuna : nomesColunas) {
            Cell cell = row.createCell(coluna++);
            cell.setCellValue(nomeColuna);
        }

        nomesColunas = this.obtemNomeColunas(linhasPlanilha);

        Class<? extends Object> clazz = linhasPlanilha.get(0).getClass();

        for (T linhaPlanilha : linhasPlanilha) {
            row = sheet.createRow(linha++);

            coluna = 0;

            for (String nomeColuna : nomesColunas) {
                Cell cell = row.createCell(coluna++);

                this.atribuiValorCelula(clazz, linhaPlanilha, nomeColuna, cell);
            }
        }
    }

    /**
     *
     */
    private List<String> obtemNomeColunas() {
        return Arrays.asList("Data de Vencimento", "Credenciador", "Participante", "Banco liquidante", "Valor emissor",
                        "Ajuste emissor", "Situação", "Data Vencimento Original", "Valor remuneração",
                        "Ajuste remuneração", "Situação", "Data Vencimento Original", "Valor líquido");
    }

    /**
     * Obtem nome das colunas
     *
     * @param linhasPlanilha
     * @return
     */
    private List<String> obtemNomeColunas(final List<T> linhasPlanilha) {
        List<String> nomesColunas = new ArrayList<>();

        try {
            Field[] campos = linhasPlanilha.get(0).getClass().getDeclaredFields();

            for (Field campo : campos) {
                nomesColunas.add(campo.getName());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return nomesColunas;
    }

    /**
     * Obtem título das colunas
     *
     * @param linhasPlanilha
     * @return
     */
    private List<String> obtemTituloColunas(final List<T> linhasPlanilha) {
        List<String> nomesColunas = new ArrayList<>();

        try {
            Field[] campos = linhasPlanilha.get(0).getClass().getDeclaredFields();

            for (Field campo : campos) {

                // if anotacao ?
                nomesColunas.add(campo.getName());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return nomesColunas;
    }

    /**
     * Prepara nome do método para invocação por reflexão
     *
     * @param atributo
     * @return
     */
    private static String preparaNomeMetodo(final String atributo) {
        if (atributo.length() == 0) {
            return atributo;
        }
        return "get" + atributo.substring(0, 1).toUpperCase() + atributo.substring(1);

    }

    /**
     * Atribui o valor do atributo na celula da planilha
     *
     * @param clazz
     * @param linhaPlanilha
     * @param nomeColuna
     * @param cell
     */
    private void atribuiValorCelula(final Class<? extends Object> clazz, final T linhaPlanilha, final String nomeColuna,
                    final Cell cell) {

        Method metodoRef = null;

        try {
            metodoRef = clazz.getMethod(preparaNomeMetodo(nomeColuna));
        } catch (NoSuchMethodException | SecurityException e) {
            metodoRef = null;
        }

        if (metodoRef != null) {
            Object valorColuna;

            try {
                valorColuna = metodoRef.invoke(linhaPlanilha, (Object[]) null);
            } catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
                valorColuna = null;
            }

            if (valorColuna != null) {
                if (valorColuna instanceof String) {
                    cell.setCellValue((String) valorColuna);
                } else if (valorColuna instanceof Integer) {
                    cell.setCellValue((Integer) valorColuna);
                } else if (valorColuna instanceof Long) {
                    cell.setCellValue((Long) valorColuna);
                } else if (valorColuna instanceof Double) {
                    cell.setCellValue((Double) valorColuna);
                }
            }
        }

    }
}
