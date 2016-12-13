/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.demo;

import java.util.List;

import com.google.common.collect.Lists;

import de.patronas.opus.opuxl.server.ExcelType;
import de.patronas.opus.opuxl.server.MatrixHeader;
import de.patronas.opus.opuxl.server.OpuxlArg;
import de.patronas.opus.opuxl.server.OpuxlMatrix;
import de.patronas.opus.opuxl.server.OpuxlMethod;

/**
 * Example Methods to register with the Opuxl Demo Server
 * @author stepan
 */
public class DemoMethods {

  /**
   * @return a list of Names
   */
  @OpuxlMethod(name = "GetNames", description = "Get some test names")
  public OpuxlMatrix getNames() {
    final OpuxlMatrix result = new OpuxlMatrix();
    final String[] names = new String[] {
        "Jon Snow", "Arya Stark", "Daenerys Targaryen", "Ramsay Bolton"
    };

    result.getData().add(Lists.newArrayList(names));
    return result;
  }

  /**
   * Create a matrix from 0 to {@linkplain end} with 10 entries per line.
   */
  @OpuxlMethod(name = "GetSeries", description = "Gets a number series for the given start and end parameter")
  public OpuxlMatrix getNumberSeries(@OpuxlArg(name = "End", description = "The Number last number of the Series.", type = ExcelType.NUMBER) final Double end) {
    final OpuxlMatrix result = new OpuxlMatrix();

    final int target = Math.abs(end.intValue());

    for (int i = 0; i < target; i++) {
      final List<Object> row = Lists.newArrayList();
      for (int j = 0; j < 10; j++) {
        row.add(10 * i + j + 1);
      }
      result.getData().add(row);
    }

    return result;
  }

  @OpuxlMethod(name = "Pow", description = "Pows the given number")
  public OpuxlMatrix pow(@OpuxlArg(name = "base", description = "The Base", type = ExcelType.NUMBER) final Double a, @OpuxlArg(name = "exponent", description = "The Exponent", type = ExcelType.NUMBER) final Double b) {
    final OpuxlMatrix result = new OpuxlMatrix();
    result.setHeaders(Lists
        .newArrayList(new MatrixHeader(a + "^" + b, ExcelType.NUMBER)));
    result.getData().add(Lists.newArrayList(Math.pow(a, b)));
    return result;
  }
}
