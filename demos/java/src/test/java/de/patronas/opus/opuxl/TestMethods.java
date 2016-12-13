/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl;

import com.google.common.collect.Lists;

import de.patronas.opus.opuxl.server.ExcelType;
import de.patronas.opus.opuxl.server.OpuxlArg;
import de.patronas.opus.opuxl.server.OpuxlMatrix;
import de.patronas.opus.opuxl.server.OpuxlMethod;

/**
 * @author stepan
 */
public class TestMethods {
  @OpuxlMethod(name = "getFoobar", description = "")
  public OpuxlMatrix getFoobar() {
    final OpuxlMatrix result = new OpuxlMatrix();
    result.getData().add(Lists.newArrayList("Foobar"));
    return result;
  }

  @OpuxlMethod(name = "square", description = "")
  public OpuxlMatrix square(@OpuxlArg(name = "Number", description = "The Number which will be squared.", type = ExcelType.NUMBER) final double number) {
    final OpuxlMatrix result = new OpuxlMatrix();
    result.getData().add(Lists.newArrayList(Math.pow(number, 2)));
    return result;
  }
}
