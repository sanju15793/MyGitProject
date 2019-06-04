import java.util.Scanner;
import java.util.regex.Pattern;

public class Regexdemo {


    public static void main(String[] args) {
        System.out.println("*********regex test**********");
        Scanner scan = new Scanner(System.in);
        System.out.println("enter any string to test the regex :");
        String exp = scan.next();

        boolean flag = Pattern.matches("([a-zA-Z0-9]*)@([a-z]+).([com])+", exp);
        String result = (flag) ? "matches" : "no match" ;
        System.out.println(result);





    }
}
