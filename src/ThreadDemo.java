public class ThreadDemo extends Thread{


    public void run()
    {
        System.out.println(Thread.currentThread().getId() );
        System.out.println(Thread.currentThread().getPriority());
    }
    public static void main(String[] args) {
        ThreadDemo td1 = new ThreadDemo();
        ThreadDemo td2 = new ThreadDemo();

        td1.setPriority(Thread.MIN_PRIORITY);
        td2.setPriority(Thread.MAX_PRIORITY);

        td1.start();
        td2.start();


    }
}
