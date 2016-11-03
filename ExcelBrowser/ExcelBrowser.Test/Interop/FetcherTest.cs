using System.Diagnostics;
using System.Threading;
using NUnit.Framework;
using System;

namespace ExcelBrowser.Interop.Test {

    [TestFixture]
    class FetcherTest {

        private static int AlwaysReturns1() => 1;
        private static int AlwaysReturns2() => 2;
        private static int Returns1Eventually() {
            Thread.Sleep(1000);
            return 1;
        }
        private static void EventShouldNotBeRaised(object sender, EventArgs e) {
            Assert.Fail("Event should not be raised.");
        }

        [Test]
        public void Fetcher_PublishesAvailableResultsQuickly() {
            var sw = new Stopwatch();
            var timeoutMs = 5;

            //Will store published result
            int? result = null;

            //Create a fetcher that fetches the constant 1
            var fetcher = new Fetcher<int>(getValue: AlwaysReturns1);
            fetcher.Fetched += (sender, e) => { result = fetcher.Result; };

            sw.Start();
            fetcher.Fetch();
            while (sw.ElapsedMilliseconds < timeoutMs && result == null) {
                //Loop until event is fired and handled, or timeout is up.
            }

            Assert.AreEqual(1, result);
        }

        [Test]
        public void Fetcher_PublishesDelayedResultsEventually() {
            var sw = new Stopwatch();
            var timeoutMs = 1100;

            //Will store published result
            int? result = null;

            //Create a fetcher that fetches the constant 1
            var fetcher = new Fetcher<int>(getValue: Returns1Eventually);
            fetcher.Fetched += (sender, e) => { result = fetcher.Result; };

            sw.Start();
            fetcher.Fetch();
            while (sw.ElapsedMilliseconds < timeoutMs && result == null) {
                //Loop until event is fired and handled, or timeout is up.
            }

            Assert.AreEqual(1, result);
            Assert.AreEqual(FetchStatus.Found, fetcher.Status);
        }

        #region Status

        [Test]
        public void Fetcher_StatusDefaultsToNotStarted() {
            var fetcher = new Fetcher<int>(AlwaysReturns1);
            Assert.AreEqual(FetchStatus.NotStarted, fetcher.Status);
        }

        [Test]
        public void Fetcher_StatusFetchingIfNoResultYet() {
            var fetcher = new Fetcher<int>(Returns1Eventually);
            fetcher.Fetch();
            Assert.AreEqual(FetchStatus.Fetching, fetcher.Status);
        }

        [Test]
        public void Fetcher_StatusFoundIfResult() {
            var sw = new Stopwatch();
            var timeoutMs = 5;

            //Will store published result
            int? result = null;

            //Create a fetcher that fetches the constant 1
            var fetcher = new Fetcher<int>(getValue: AlwaysReturns1);
            fetcher.Fetched += (sender, e) => { result = fetcher.Result; };

            sw.Start();
            fetcher.Fetch();
            while (sw.ElapsedMilliseconds < timeoutMs && result == null) {
                //Loop until event is fired and handled, or timeout is up.
            }
            
            Assert.AreEqual(FetchStatus.Found, fetcher.Status);
        }

        [Test]
        public void Fetcher_StatusErrorIfUnfilteredExceptionThrown() {
            var sw = new Stopwatch();
            var timeoutMs = 5;

            var fetcher = new Fetcher<int>(
                getValue: AlwaysThrowsNullReferenceException,
                exceptionFilter: FilterAllButNullReferenceExceptions);

            fetcher.Fetched += EventShouldNotBeRaised;

            sw.Start();
            fetcher.Fetch();
            while (sw.ElapsedMilliseconds < timeoutMs) {
                //Loop until event is fired and handled, or timeout is up.
            }

            Assert.AreEqual(FetchStatus.Error, fetcher.Status);
        }

        #endregion
        
        #region Exception filter

        private int AlwaysThrowsNullReferenceException() {
            string x = null;
            return x.Length;
        }

        private int AlwaysThrowsDivideByZeroException() {
            int x = 0;
            return 1 / x;
        }

        private bool FilterAllButNullReferenceExceptions(Exception x) {
            return x.GetType() != typeof(NullReferenceException);
        }

        [Test]
        public void Fetcher_KeepsTryingWhenExceptionThrownByDefault() {
            var sw = new Stopwatch();
            var timeoutMs = 5;

            var fetcher = new Fetcher<int>(getValue: AlwaysThrowsDivideByZeroException);
            fetcher.Fetched += EventShouldNotBeRaised;

            sw.Start();
            fetcher.Fetch();
            while (sw.ElapsedMilliseconds < timeoutMs) {
                //Loop until event is fired and handled, or timeout is up.
            }

            Assert.AreEqual(FetchStatus.Fetching, fetcher.Status);
            Assert.Throws<InvalidOperationException>(() => { var x = fetcher.Result; });
        }

        [Test]
        public void Fetcher_ExceptionFilterKeepsTryingIfExceptionReturnsTrue() {
            var sw = new Stopwatch();
            var timeoutMs = 5;

            var fetcher = new Fetcher<int>(
                getValue: AlwaysThrowsDivideByZeroException,
                exceptionFilter: FilterAllButNullReferenceExceptions);

            fetcher.Fetched += EventShouldNotBeRaised;

            sw.Start();
            fetcher.Fetch();
            while (sw.ElapsedMilliseconds < timeoutMs) {
                //Loop until event is fired and handled, or timeout is up.
            }

            Assert.AreEqual(FetchStatus.Fetching, fetcher.Status);
            Assert.Throws<InvalidOperationException>(() => { var x = fetcher.Result; });
        }

        [Test]
        public void Fetcher_ExceptionFilterThrowsIfExceptionReturnsFalse() {
            var sw = new Stopwatch();
            var timeoutMs = 5;

            var fetcher = new Fetcher<int>(
                getValue: AlwaysThrowsNullReferenceException,
                exceptionFilter: FilterAllButNullReferenceExceptions);

            fetcher.Fetched += (sender, e) => { Assert.Fail("Event should not be raised."); };

            sw.Start();
            fetcher.Fetch();
            while (sw.ElapsedMilliseconds < timeoutMs) {
                //Loop until event is fired and handled, or timeout is up.
            }

            Assert.AreEqual(FetchStatus.Error, fetcher.Status);
            Assert.Throws<InvalidOperationException>(() => { var x = fetcher.Result; });
        }

        #endregion

        #region Value filter

        private static bool FilterEvenValues(int x) {
            return x % 2 == 0;
        }

        [Test]
        public void Fetcher_ValueFilterPreventsValuesThatReturnFalseFromBeingFound() {
            var sw = new Stopwatch();
            var timeoutMs = 5;

            var fetcher = new Fetcher<int>(
                getValue: AlwaysReturns1,
                valueFilter: FilterEvenValues);

            fetcher.Fetched += EventShouldNotBeRaised;

            sw.Start();
            fetcher.Fetch();
            while (sw.ElapsedMilliseconds < timeoutMs) {
                //Loop until event is fired and handled, or timeout is up.
            }

            Assert.AreEqual(FetchStatus.Fetching, fetcher.Status);
            Assert.Throws<InvalidOperationException>(() => { var x = fetcher.Result; });
        }

        [Test]
        public void Fetcher_ValueFilterAllowsValuesThatReturnTrueToBeFound() {
            var sw = new Stopwatch();
            var timeoutMs = 5;

            //Will store published result
            int? result = null;

            //Create a fetcher that fetches the constant 1
            var fetcher = new Fetcher<int>(getValue: AlwaysReturns2);
            fetcher.Fetched += (sender, e) => { result = fetcher.Result; };

            sw.Start();
            fetcher.Fetch();
            while (sw.ElapsedMilliseconds < timeoutMs && result == null) {
                //Loop until event is fired and handled, or timeout is up.
            }

            Assert.AreEqual(2, result);
            Assert.AreEqual(FetchStatus.Found, fetcher.Status);
        }

        #endregion
    }
}
