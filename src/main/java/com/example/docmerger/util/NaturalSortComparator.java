package com.example.docmerger.util;

import com.example.docmerger.model.FileItem;

import java.util.Comparator;

public final class NaturalSortComparator implements Comparator<FileItem> {
    @Override
    public int compare(FileItem left, FileItem right) {
        return naturalCompare(left.getFileName(), right.getFileName());
    }

    private static int naturalCompare(String left, String right) {
        int leftIndex = 0;
        int rightIndex = 0;
        int leftLength = left.length();
        int rightLength = right.length();

        while (leftIndex < leftLength && rightIndex < rightLength) {
            char leftChar = left.charAt(leftIndex);
            char rightChar = right.charAt(rightIndex);

            if (Character.isDigit(leftChar) && Character.isDigit(rightChar)) {
                int leftStart = leftIndex;
                int rightStart = rightIndex;

                while (leftIndex < leftLength && Character.isDigit(left.charAt(leftIndex))) {
                    leftIndex++;
                }
                while (rightIndex < rightLength && Character.isDigit(right.charAt(rightIndex))) {
                    rightIndex++;
                }

                String leftNumber = trimLeadingZeros(left.substring(leftStart, leftIndex));
                String rightNumber = trimLeadingZeros(right.substring(rightStart, rightIndex));

                int lengthCompare = Integer.compare(leftNumber.length(), rightNumber.length());
                if (lengthCompare != 0) {
                    return lengthCompare;
                }

                int numberCompare = leftNumber.compareTo(rightNumber);
                if (numberCompare != 0) {
                    return numberCompare;
                }
            } else {
                int compare = Character.compare(Character.toLowerCase(leftChar), Character.toLowerCase(rightChar));
                if (compare != 0) {
                    return compare;
                }
                leftIndex++;
                rightIndex++;
            }
        }

        return Integer.compare(leftLength - leftIndex, rightLength - rightIndex);
    }

    private static String trimLeadingZeros(String value) {
        String trimmed = value.replaceFirst("^0+", "");
        return trimmed.isEmpty() ? "0" : trimmed;
    }
}
