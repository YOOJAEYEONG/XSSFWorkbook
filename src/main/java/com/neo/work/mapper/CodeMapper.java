package com.neo.work.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.apache.ibatis.annotations.Select;
import org.springframework.lang.Nullable;

import java.util.List;
import java.util.Map;


@Mapper
public interface CodeMapper {

  @Select("SELECT LEA_DONG_CD, SIDO, SIGUNGU, EUPMYUNDONG FROM TB_ADDR_INFO WHERE EUPMYUNDONG IS NULL AND SIGUNGU IS NOT NULL")
  List<Map> selectSigungu();

  @Select("SELECT cd, cd_nm, cd_grp, expnt FROM TB_COMM_CD WHERE CD_GRP = #{cd_grp}")
  List<Map> selectCodeGroup(@Nullable @Param("cd_grp")String cd_grp);

}
