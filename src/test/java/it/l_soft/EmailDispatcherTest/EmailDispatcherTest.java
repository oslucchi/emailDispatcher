package it.l_soft.EmailDispatcherTest;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class EmailDispatcherTest 
    extends TestCase
{
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public EmailDispatcherTest( String testName )
    {
        super( testName );
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite()
    {
        return new TestSuite( EmailDispatcherTest.class );
    }

    /**
     * Rigourous Test :-)
     */
    public void testEmailDispatcher()
    {
        assertTrue( true );
    }
}
