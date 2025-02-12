const tester = TestGAS.createExecutor();


class Test_aggregate {
  test_aggregate_no_events() {
    const prefixes = ['prefix1', 'prefix2'];
    const events = [];
    const wholeDate = ['2024-01-01', '2024-01-02'];
    const result = aggregate(prefixes, events, wholeDate);

    tester.assertEquals(Object.keys(result.prefixDurations).length, 0, "prefixDurations should be empty");
    tester.assertEquals(Object.keys(result.prefixDailyDurations).length, 2, "prefixDailyDurations should have 2 entries");
    tester.assertEquals(result.prefixDailyDurations['2024-01-01']['prefix1'], 0, "prefixDailyDurations should be 0 for prefix1 on 2024-01-01");
  }

  // 他のテストケースもここに追加
  test_aggregate_with_events() {
    const prefixes = ['prefix1', 'prefix2'];
    const wholeDate = ['2024-01-01', '2024-01-02'];
    const mockEvent = {
      getTitle: () => 'prefix1 Event',
      getStartTime: () => new Date('2024-01-01T10:00:00'),
      getEndTime: () => new Date('2024-01-01T11:00:00'),
    };
    const events = [mockEvent];
    const result = aggregate(prefixes, events, wholeDate);

    tester.assertEquals(result.prefixDurations['prefix1'], 3600000, "prefixDurations should be 3600000");
    tester.assertEquals(result.prefixDailyDurations['2024-01-01']['prefix1'], 3600000, "prefixDailyDurations should be 3600000");
  }
}

function runAllTests() {
  let failureFuncs = tester.executeTestGas(Test_aggregate);
  if (failureFuncs.length > 0) {
    Logger.log("テスト失敗: " + failureFuncs.join(", "));
  } else {
    Logger.log("すべてのテストが成功しました。");
  }
}