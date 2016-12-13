/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Marks a method discoverable for the Opuxl Plugin. The return type of a method
 * should always be an OpuxlMatrix.
 * @author stepan
 */
@Target(value = ElementType.METHOD)
@Retention(value = RetentionPolicy.RUNTIME)
public @interface OpuxlMethod {
  /**
   * (optional) The Name of this registered method.
   */
  String name() default "";

  String description();
}
