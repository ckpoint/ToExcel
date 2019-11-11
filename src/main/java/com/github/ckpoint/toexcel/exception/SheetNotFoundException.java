package com.github.ckpoint.toexcel.exception;

/**
 * The type Sheet not found exception.
 */
public class SheetNotFoundException extends RuntimeException {

    /**
     * Instantiates a new Sheet not found exception.
     *
     * @param message the message
     */
    public SheetNotFoundException(String message) {
        super(message);
    }
}
