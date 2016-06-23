/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl;

import java.lang.annotation.Annotation;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.collect.Lists;

/**
 * The function handler is responsible to store the registered functions and to
 * execute them.
 * @author stepan
 */
public class FunctionHandler {

  private static final Logger LOG = LoggerFactory
      .getLogger(FunctionHandler.class);

  private final Map<String, InstanceMethod> methods = new HashMap<>();

  public void registerMethods(final Object instance) {
    final Method[] instanceMethods = instance.getClass().getMethods();
    // Iterate over the methods of the given object instance and register all
    // 'Opuxl' marked methods.
    for (final Method method : instanceMethods) {
      final OpuxlMethod annotation = method.getAnnotation(OpuxlMethod.class);
      if (annotation != null) {

        final List<ExcelArgumentAttribute> parameterArguments = Lists
            .newArrayList();

        final Annotation[][] parameterAnnotations = method
            .getParameterAnnotations();
        for (final Annotation[] annotations : parameterAnnotations) {
          for (final Annotation paramAnnotation : annotations) {

            if (paramAnnotation.annotationType().equals(OpuxlArg.class)) {
              parameterArguments
                  .add(new ExcelArgumentAttribute((OpuxlArg) paramAnnotation));
            }
          }
        }

        final String registeredName = annotation.name().isEmpty() ? method
            .getName() : annotation.name();
        getMethods()
            .put(registeredName, new InstanceMethod(instance, method, parameterArguments));
      }
    }
  }

  public Object[][] execute(final FunctionPayload payload) {
    if (!hasFunction(payload.getName())) {
      throw new NoSuchMethodError("No method found with name: "
          + payload.getName());
    }
    return getMethods().get(payload.getName()).execute(payload.getArgs());
  }

  public boolean hasFunction(final String name) {
    return getMethods().containsKey(name);
  }

  public Map<String, InstanceMethod> getMethods() {
    return methods;
  }

}
