//#######################################################################################
#include "stdafx.h"
#ifdef _TEST  
#include "CppUnitTest.h"
#include "test.h" 
using namespace Microsoft::VisualStudio::CppUnitTestFramework;
int Foo(int a, int b){
	if (a == 0 || b == 0)
	{
		throw "don't do that";
	}
	int c = a % b;
	if (c == 0)
		return b;
	return Foo(b, c);
}



namespace MyTest {


	TEST_CLASS(MyTests){
		public:
			TEST_METHOD(MyTestMethod){
			Assert::AreEqual(Foo(2, 3), 1);     
			}
			TEST_METHOD(MyTestMethod_2){
		    Assert::AreEqual(Foo(2, 3),1);       
			} 
		};
	



}
#endif 