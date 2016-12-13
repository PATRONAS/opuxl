/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Describes a parameter of a discoverable Method.
 * @author stepan
 */
@Target(value = ElementType.PARAMETER)
@Retention(value = RetentionPolicy.RUNTIME)
public @interface OpuxlArg {
  /**
   * The Name of the Argument.
   */
  String name();

  /**
   * The description of the Argument.
   */
  String description();

  /**
   * The Excel Type of the Argument.
   */
  ExcelType type();

  /**
   * Whether the Argument is optional.
   */
  boolean optional() default false;
}
