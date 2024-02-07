package com.impllife.xlsx.service.res;

public class ResPathGetter {
    public static String getResPath(String resName) {
        return ResPathGetter.class.getClassLoader().getResource(resName).getPath();
    }
}
