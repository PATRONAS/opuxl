/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.demo;

import de.patronas.opus.opuxl.ExcelType;
import de.patronas.opus.opuxl.OpuxlArg;
import de.patronas.opus.opuxl.OpuxlMethod;

/**
 * Example Methods to register with the Opuxl Demo Server
 * @author stepan
 */
public class DemoMethods {

  /**
   * @return a list of Names
   */
  @OpuxlMethod
  public String[] getNames() {
    return new String[] {
        "Jon Snow", "Arya Stark", "Daenerys Targaryen", "Ramsay Bolton"
    };
  }

  /**
   * Create a matrix from 0 to {@linkplain end} with 10 entries per line.
   */
  @OpuxlMethod
  public Object[][] getNumberSeries(@OpuxlArg(name = "End", description = "The Number last number of the Series.", type = ExcelType.NUMBER) final Double end) {
    final int target = Math.abs(end.intValue());
    final Object[][] result = new Object[1 + target / 10][10];

    int start = 0;
    while (start < end) {
      result[start / 10][start % 10] = start++;
    }

    return result;
  }

  @OpuxlMethod
  public Object[][] pow(@OpuxlArg(name = "base", description = "The Base", type = ExcelType.NUMBER) final Double a, @OpuxlArg(name = "exponent", description = "The Exponent", type = ExcelType.NUMBER) final Double b) {
    return new Object[][] {
      {
        Math.pow(a, b)
      }
    };
  }

}
