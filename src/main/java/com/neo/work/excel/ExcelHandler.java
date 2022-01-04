package com.neo.work.excel;

import com.neo.work.mapper.CodeMapper;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;

@Slf4j
@Component
public class ExcelHandler implements ApplicationRunner {

  @Autowired
  CodeMapper codeMapper;

  List<Map> sido = null;
  List<Map> sigungu = null;
  List<Map> 검사결과코드 = null;
  List<Map> 계량기구분코드 = null;
  List<Map> 판수동저울정확도등급 = null;
  List<String> columnName = new ArrayList<>();
  Set<Integer> invalidateCell = new HashSet<>();
  HashMap<Integer, Set<Integer>> invalidateColumn = new HashMap<>();
  //수정할 엑셀파일명을 입력한다.
  String filename = "정기검사관리 (21년)"+".xlsx";


  @Override
  public void run(ApplicationArguments args)  {
    try ( BufferedInputStream file =
            new BufferedInputStream(
              new FileInputStream("C:/Users/YooJaeYeong/Desktop/excel/"+filename));
          XSSFWorkbook wb = new XSSFWorkbook(file);
          ){
      readExcelSheet(wb);
    }catch (Exception e){
      log.error(e.getMessage(),e);
    }
  }

  private void readExcelSheet(XSSFWorkbook wb) {
    try {
      sido = codeMapper.selectSigungu();
//      sido.forEach(map -> log.debug("{}",map.values()));
      //[CD, EXPNT, CD_GRP, CD_NM]
      검사결과코드 = codeMapper.selectCodeGroup("AC032000");
      검사결과코드.forEach(map -> log.debug("{}",map.values()));
      계량기구분코드 = codeMapper.selectCodeGroup("AB001000");
      계량기구분코드.forEach(map -> log.debug("{}",map.values()));
      판수동저울정확도등급 = codeMapper.selectCodeGroup("AA101000");
      판수동저울정확도등급.forEach(map -> log.debug("{}",map.values()));

      CellStyle headerStyle = defineCellStyle(wb.createCellStyle(),"header");
      CellStyle errorStyle = defineCellStyle(wb.createCellStyle(),"error");
      CellStyle defaultStyle = defineCellStyle(wb.createCellStyle(),"default");


      //시트 수 (첫번째에만 존재하므로 0을 준다)
      //만약 각 시트를 읽기위해서는 FOR 문을 한번더 돌려준다
      Sheet sheet = wb.getSheetAt(0);
      sheet.autoSizeColumn(400,true);
      //컬럼을 추가하기위함
      Row row0 = sheet.getRow(0);
      CellStyle cellStyle = sheet.getRow(0).getCell(0).getCellStyle();
      //컬럼명을 추출
      row0.forEach(cell -> columnName.add(cell.getStringCellValue()));
      log.warn("\n변경전 columnName {}",columnName);


      //법정동코드 컬럼 추가
      row0 = sheet.getRow(0);
      row0.shiftCellsRight(getColumnIndex("검사결과",sheet),row0.getLastCellNum(),1);
      Cell cell_LEA_DONG_CD = row0.createCell(getColumnIndex("검사결과",sheet)-1, CellType.STRING);
      cell_LEA_DONG_CD.setCellValue("LEA_DONG_CD");

      //검사결과코드 컬럼추가
      row0 = sheet.getRow(0);
      row0.shiftCellsRight(getColumnIndex("불합격사유",sheet),row0.getLastCellNum(),1);
      Cell cell_CHEK_RELT_CD = row0.createCell(getColumnIndex("불합격사유",sheet)-1, CellType.STRING);
      cell_CHEK_RELT_CD.setCellValue("CHEK_RELT_CD");

      //저울종류코드 컬럼추가
      row0 = sheet.getRow(0);
      row0.shiftCellsRight(getColumnIndex("기물번호",sheet),row0.getLastCellNum(),1);
      Cell cell_SCAE_KND_CD = row0.createCell(getColumnIndex("기물번호",sheet)-1, CellType.STRING);
      cell_SCAE_KND_CD.setCellValue("SCAE_KND_CD");

      //등급코드 컬럼추가
      row0 = sheet.getRow(0);
      row0.shiftCellsRight(getColumnIndex("등록일",sheet),row0.getLastCellNum(),1);
      Cell cell_RATG_CD = row0.createCell(getColumnIndex("등록일",sheet)-1, CellType.STRING);
      cell_RATG_CD.setCellValue("RATG_CD");


      columnName.clear();
      row0.forEach(cell -> columnName.add(cell.getStringCellValue()));
      log.warn("\n변경후 columnName {}",columnName);
      log.info("엑셀 수정 시작..");

      for(int rowindex = 0; rowindex <= row0.getSheet().getLastRowNum(); rowindex++){
        //시군구값으로 찾은 법정동(리동) 코드값을 임시 저장함.
        List<Object> lea_dong_cd = new ArrayList<>();
        //값으로 찾은 코드값을 저장함
        Map<String,String> codeValues = new HashMap<>();
        Row row = sheet.getRow(rowindex);
        //디버깅을위한 값 출력용
        final StringBuffer value =  new StringBuffer();

        for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
          Cell cell = row.getCell(cellIndex);
          cell.setCellStyle(defineCellStyle(cell.getCellStyle(),"default"));
          //첫줄은 제목컬럼이므로 관련 스타일을 지정
          if (cell.getRowIndex() == 0){
            cell.setCellStyle(headerStyle);
            row.setHeight((short)600);
          }
          switch (cell.getCellType()){
            case FORMULA:
              value.append(cell.getCellFormula()+"\t");
              log.warn("FORMULA : {}",cell.getCellFormula());
              break;
            case NUMERIC:
              if (cell.getColumnIndex() != 0)
                log.warn("NUMERRIC : {}",(int)cell.getNumericCellValue());
              value.append((int)cell.getNumericCellValue()+"\t");
              break;
            case STRING:
            case BLANK:
            case _NONE:
              //컬럼명을 제외함
              if(cell.getRowIndex() > 0){
                //셀에 빈값, 공백값이 존재하는경우
                if (cell.getStringCellValue().trim().isBlank() || cell.getStringCellValue().trim().isEmpty()){
                  //셀이 notNull 조건인지
                  if (isNotNullColumn(cell)){
                    addInvalidCell(cell);
                    cell.setCellStyle(errorStyle);
                    break;
                  }
                }

                switch (getColumnValue(cell)){
                  case "시도" :
                    //sigungu 코드를 추가하기위해 sido를 필터링함
                    sigungu = sido.stream()
                      .filter(map -> map.get("SIDO").equals(cell.getStringCellValue()))
                      .collect(Collectors.toList());
                    break;
                  case "시군구" :
                    lea_dong_cd = sigungu.stream()
                      .filter(map -> map.get("SIGUNGU").equals(cell.getStringCellValue()))
                      .map(e -> e.get("LEA_DONG_CD"))
                      .collect(Collectors.toList());
                    if(lea_dong_cd.size() > 1)
                      log.error("sigungu 필터갯수가 하나 이상 발견되는 문제 발생 : {}",lea_dong_cd);
                    if(lea_dong_cd.size() == 0){
                      log.warn("sigungu 필터결과가 없음");
                      cell.setCellStyle(errorStyle);
                    }
                    if (lea_dong_cd.size() == 1){
                      codeValues.put("LEA_DONG_CD", (String) lea_dong_cd.get(0));
                      value.append(cell.getStringCellValue()+"\t");
                    }
                    break;
                  case "LEA_DONG_CD" ://row0 에서 추가한컬럼
                    row.shiftCellsRight(cell.getColumnIndex(),row.getLastCellNum(),1);
                    row.createCell(cell.getColumnIndex()-1, CellType.STRING)
                      .setCellValue(codeValues.get("LEA_DONG_CD"));
                    value.append(codeValues.get("LEA_DONG_CD")+"\t");
                    lea_dong_cd.clear();
                    break;
                  case "검사결과" :
                    List<Object> chek_relt_cd = 검사결과코드.stream()
                      .filter(map -> map.get("CD_NM").equals(cell.getStringCellValue().trim()))
                      .map(e -> e.get("CD"))
                      .collect(Collectors.toList());

                    if(chek_relt_cd.size() > 1)
                      log.error("검사결과코드 필터갯수가 하나 이상 발견됨 : {}",chek_relt_cd.size());
                    if(chek_relt_cd.size() == 0) {
                      log.error("검사결과코드 필터결과가 없음");
                      cell.setCellStyle(errorStyle);
                    }
                    if(chek_relt_cd.size() == 1){
                      value.append(cell.getStringCellValue()+"\t");
                      codeValues.put("CHEK_RELT_CD", (String) chek_relt_cd.get(0));
                    }
                    break;
                  case "불합격사유":
                    //합격인데 불합격사유가 존재 하는경우
                    if (row.getCell(getColumnIndex("검사결과",sheet)).getStringCellValue().trim().equals("합격") &&
                      !row.getCell(getColumnIndex("불합격사유",sheet)).getStringCellValue().trim().isBlank() &&
                      !row.getCell(getColumnIndex("불합격사유",sheet)).getStringCellValue().trim().isEmpty()){
                      log.warn("합격인데 불합격사유가 존재 하는경우 {}행 {}열",cell.getRowIndex(),cell.getColumnIndex());
                      addInvalidCell(cell);
                      cell.setCellStyle(errorStyle);
                    }
                    value.append(cell.getStringCellValue()+"\t");
                    break;
                  case "CHEK_RELT_CD" ://검사결과코드
                    row.shiftCellsRight(cell.getColumnIndex(),row.getLastCellNum(),1);
                    row.createCell(cell.getColumnIndex()-1, CellType.STRING)
                      .setCellValue(codeValues.get("CHEK_RELT_CD"));
                    value.append(codeValues.get("CHEK_RELT_CD")+"\t");
                    break;
                  case "저울종류" :
                    List<Object> scae_knd_cd = 계량기구분코드.stream()
                      .filter(map -> map.get("CD_NM").equals(cell.getStringCellValue().trim()))
                      .map(e -> e.get("CD"))
                      .collect(Collectors.toList());
                    if(scae_knd_cd.size() > 1)
                      log.error("계량기구분코드 필터갯수가 하나 이상 발견됨 : {}",scae_knd_cd.size());
                    if(scae_knd_cd.size() == 0){
                      log.error("계량기구분코드 필터결과가 없음");
                      cell.setCellStyle(errorStyle);
                    }
                    if(scae_knd_cd.size() == 1){
                      value.append(cell.getStringCellValue()+"\t");
                      codeValues.put("SCAE_KND_CD", (String) scae_knd_cd.get(0));
                    }
                    break;
                  case "SCAE_KND_CD" :// 저울종류코드값
                    row.shiftCellsRight(cell.getColumnIndex(),row.getLastCellNum(),1);
                    row.createCell(cell.getColumnIndex()-1, CellType.STRING)
                      .setCellValue(codeValues.get("SCAE_KND_CD"));
                    cell.setCellStyle(defaultStyle);
                    value.append(codeValues.get("SCAE_KND_CD")+"\t");
                    break;
                  case "등급" :
                    if(cell.getStringCellValue().trim().isBlank()) break;
                    List<Object> ratg_cd = 판수동저울정확도등급.stream()
                      .filter(map -> map.get("CD_NM").equals(cell.getStringCellValue()))
                      .map(e -> e.get("CD"))
                      .collect(Collectors.toList());
                    if(ratg_cd.size() > 1)
                      log.error("판수동저울정확도등급 필터갯수가 하나 이상 발견됨 : {}",ratg_cd.size());
                    if(ratg_cd.size() == 0) {
                      log.error("판수동저울정확도등급 필터결과가 없음");
                      cell.setCellStyle(errorStyle);
                    }
                    if(ratg_cd.size() == 1){
                      value.append(ratg_cd.get(0)+"\t");
                      codeValues.put("RATG_CD", (String) ratg_cd.get(0));
                    }
                    break;
                  case "RATG_CD":
                    value.append(codeValues.get("RATG_CD")+"\t");
                    row.shiftCellsRight(cell.getColumnIndex(),row.getLastCellNum(),1);
                    row.createCell(cell.getColumnIndex()-1, CellType.STRING)
                      .setCellValue(codeValues.get("RATG_CD"));
                    break;
                  default:
                    value.append(cell.getStringCellValue());
                    if (cell.getColumnIndex() == row.getLastCellNum())
                      value.append("\n");
                    else
                      value.append("\t");
                    if (isNotNullColumn(cell)){
                      addInvalidCell(cell);
                      cell.setCellStyle(errorStyle);
                    }
                }
              }
              break;
            case BOOLEAN:
              log.warn("BOOLEAN : {}",cell.getBooleanCellValue());
              break;
            case ERROR:
              log.error("ERROR : {}",cell.getErrorCellValue());
              break;
            default:
              log.error("알려지지않은 타입");
              break;
          }
          //로그 확인
          if (cell.getRowIndex() % 1000 == 0 && cell.getRowIndex()!= 0 &&
            cell.getColumnIndex()+1 == cell.getRow().getLastCellNum()){
            log.info("{}번째 행 작업중",cell.getRowIndex());
          }
        }
      }
      writeExcelSheet(wb);
    }catch (Exception e){
      log.info(e.getMessage(),e);
    }
  }

  /**
   * 엑셀 저장
   * @param wb
   */
  private void writeExcelSheet(XSSFWorkbook wb) {
    try (BufferedOutputStream outputStream =
           new BufferedOutputStream(
             new FileOutputStream("C:/Users/YooJaeYeong/Desktop/excel2/"+filename+".xlsx"));){
      log.warn("엑셀저장중>>>>>");
      wb.write(outputStream);
    }catch (Exception e) {
      log.error(e.getMessage(),e);
    }finally {
      log.info("완료");
    }
  }


  /**
   * 올바은 값이 아닌 셀이 발견되면 따로 저장해둔다.
   * @param cell
   */
  private void addInvalidCell(Cell cell ) {
    invalidateCell = invalidateColumn.get(cell.getRowIndex());
    if (invalidateCell == null) {
      invalidateCell = new HashSet<>();
      invalidateCell.add(cell.getColumnIndex());
    }else{
      invalidateCell.add(cell.getColumnIndex());
    }
    invalidateColumn.put(cell.getRowIndex(),invalidateCell);
  }

  /**
   * 현재 cell 이 NotNull 로 지정되어있으면 true
   * 0 : NO
   * 1 : 검사기관
   * 2 : 검사일
   * 3 : 시도
   * 4 : 시군구
   * 5 : LEA_DONG_CD
   * 6 : 검사결과
   * 7 : CHEK_RELT_CD
   * 8 : 불합격사유
   * 9 : 기타
   * 10: 상호명
   * 11: 대표자
   * 12: 전화번호
   * 13: 사업장주소
   * 14: 저울종류
   * 15: SCAE_KND_CD
   * 16: 기물번호
   * 17: 저울제조사
   * 18: 구입년도
   * 19: 최대용량
   * 20: 검정눈금
   * 21: 등급
   * 22: RATG_CD
   * 23: 등록일
   * @param cell
   * @return
   */
  private boolean isNotNullColumn(Cell cell){
    //not Null 로 지정할 컬럼 인덱스
    int[] notNullColumn = {3,4,6,14,21};
    return Arrays.stream(notNullColumn).anyMatch(idx -> idx == cell.getColumnIndex());
  }

  /**
   * 현재 cell 의 컬럼이름을 리턴
   * @param cell
   * @return
   */
  private String getColumnValue(Cell cell){
    Row row = cell.getRow().getSheet().getRow(0);
    AtomicReference<String> columnName = new AtomicReference<>("");
    row.forEach(cell1 -> {
        if (cell1.getColumnIndex() == cell.getColumnIndex())
          columnName.set(cell1.getStringCellValue().trim());
    });
    log.debug("getColumnValue:return:{}",columnName);
    return columnName.get();
  }

  /**
   * 현재 cell 의 columnIndex 를 반환
   * @param column
   * @param sheet
   * @return
   */
  private int getColumnIndex(String column , Sheet sheet){
    AtomicInteger rowIndex = new AtomicInteger();
    sheet.getRow(0).forEach(cell -> {
      if(cell.getStringCellValue().trim().equals(column)){
        rowIndex.set(cell.getColumnIndex());
        return;
      }
    });
    return rowIndex.get();
  }

  /**
   * 추가한 셀에대한 스타일을 정의
   * @param style
   * @return
   */
  private CellStyle defineCellStyle(CellStyle style,String option){
    if (option.equals("error")){
      style.setFillForegroundColor(IndexedColors.RED.getIndex());
    }else if (option.equals("header")){
      style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
      style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
      style.setAlignment(HorizontalAlignment.CENTER);
      style.setVerticalAlignment(VerticalAlignment.CENTER);
    }else if (option.equals("default")){
      style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
    }
    style.setBorderTop(BorderStyle.THIN);
    style.setBorderBottom(BorderStyle.THIN);
    style.setBorderLeft(BorderStyle.THIN);
    style.setBorderRight(BorderStyle.THIN);
    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
    style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
    style.setRightBorderColor(IndexedColors.BLACK.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    return style;
  }
}
