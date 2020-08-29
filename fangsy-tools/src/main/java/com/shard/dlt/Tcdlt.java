package com.shard.dlt;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

/**
 * 体彩大乐透随机生成彩票
 *
 * @Author : fsy
 * @Date: 2020-07-28 21:13
 */
public class Tcdlt {

    public static void main(String[] args) {
        createNum();
        createNum();
        createNum();
    }


    public static void createNum() {
        Set<Integer> redNum = new HashSet<Integer>();
        Set<Integer> blueNum = new HashSet<Integer>();

        while (true) {
            if (redNum.size() >= 5) {
                break;
            }
            int num = (int) (Math.random() * 35) + 1;
            redNum.add(num);
        }

        while (true) {
            if (blueNum.size() >= 2) {
                break;
            }
            int num = (int) (Math.random() * 12) + 1;
            blueNum.add(num);
        }

        Object[] red = redNum.toArray();
        Arrays.sort(red);
        for (int i = 0; i < red.length; i++) {
            System.out.print(red[i] + "\t");
        }
        System.out.print("┼" + "\t");
        Object[] blue = blueNum.toArray();
        Arrays.sort(blue);
        for (int i = 0; i < blue.length; i++) {
            System.out.print(blue[i] + "\t");
        }
        System.out.print("\n");
    }

}
