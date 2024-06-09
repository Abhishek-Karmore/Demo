package Demo;

class program
{

    void display(int a)
    {
        int f=1;
        for(int i=2;i<=a;i++)
        {
            f=f*i;
        }
          System.out.println(f);
    }
}
public class Fact {
    public static void main(String [] args)

    {
        program obj=new program();
        obj.display(5);
    }

}

