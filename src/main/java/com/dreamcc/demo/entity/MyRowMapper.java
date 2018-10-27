package com.dreamcc.demo.entity;

import org.springframework.jdbc.core.RowMapper;

import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * @Title: demo
 * @Package: com.dreamcc.demo.entity
 * @Description:
 * @Author: dreamcc
 * @Date: 2018-10-27 08:19
 * @Version: V1.0
 */
public class MyRowMapper implements RowMapper<IOCEntity> {
	@Override
	public IOCEntity mapRow(ResultSet rs, int rowNum) throws SQLException {
		String dept = rs.getString("dept");
		String resource = rs.getString("resource");
		String info = rs.getString("info");

		IOCEntity iocEntity = new IOCEntity();
		iocEntity.setDept(dept);
		iocEntity.setResource(resource);
		iocEntity.setInfo(info);
		return iocEntity;
	}
}
