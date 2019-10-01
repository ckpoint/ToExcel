package com.github.ckpoint.toexcel.util;

import org.modelmapper.ModelMapper;
import org.modelmapper.config.Configuration;
import org.modelmapper.convention.MatchingStrategies;

/**
 * The type Model mapper generator.
 */
public class ModelMapperGenerator {

    /**
     * Default model mapper model mapper.
     *
     * @return the model mapper
     */
    public static ModelMapper defaultModelMapper() {
        return new ModelMapper();
    }

    /**
     * Strict model mapper model mapper.
     *
     * @return the model mapper
     */
    public static ModelMapper strictModelMapper() {
        ModelMapper modelMapper = defaultModelMapper();
        modelMapper.getConfiguration().setMatchingStrategy(MatchingStrategies.STRICT);
        return modelMapper;
    }

    /**
     * Enable field model mapper model mapper.
     *
     * @param accessLevel the access level
     * @return the model mapper
     */
    public static ModelMapper enableFieldModelMapper(Configuration.AccessLevel accessLevel) {
        ModelMapper modelMapper = strictModelMapper();
        modelMapper.getConfiguration().setFieldMatchingEnabled(true);
        modelMapper.getConfiguration().setFieldAccessLevel(accessLevel);
        return modelMapper;
    }

    /**
     * Enable field model mapper model mapper.
     *
     * @return the model mapper
     */
    public static ModelMapper enableFieldModelMapper() {
        return enableFieldModelMapper(Configuration.AccessLevel.PRIVATE);
    }
}
