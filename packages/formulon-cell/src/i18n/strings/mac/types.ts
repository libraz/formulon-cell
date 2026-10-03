export interface MacStrings {
  macData: {
    goalSeek: {
      title: string;
      formulaCell: string;
      targetValue: string;
      changingCell: string;
      iterations: string;
      run: string;
      cancel: string;
      result: string;
      invalidCell: string;
      invalidTarget: string;
      calculating: string;
      converged: string;
    };
    consolidate: {
      title: string;
      sources: string;
      sourcesHint: string;
      destination: string;
      function: string;
      replace: string;
      labels: string;
      links: string;
      unsupportedOption: string;
      run: string;
      cancel: string;
    };
    subtotal: {
      title: string;
      range: string;
      groupBy: string;
      columns: string;
      function: string;
      replace: string;
      summaryBelow: string;
      unsupportedOption: string;
      run: string;
      cancel: string;
    };
    functions: { sum: string; average: string; count: string; min: string; max: string };
    /** Keyed by the error codes returned from the Data command planners. */
    errors: {
      invalidRange: string;
      invalidColumn: string;
      targetNotFinite: string;
      sameCell: string;
      noFormula: string;
      changingHasFormula: string;
      changingNotNumeric: string;
      scratchUnsupported: string;
      invalidTolerance: string;
      invalidIterations: string;
      nonFiniteResult: string;
      notConverged: string;
      stale: string;
      notEditable: string;
      writeFailed: string;
      failed: string;
      noSources: string;
      invalidFunction: string;
      invalidSource: string;
      areaLimit: string;
      dimensionMismatch: string;
      invalidDestination: string;
      destinationOverlapsSource: string;
      activeSheetOnly: string;
      needsDataRow: string;
      engineUnsupported: string;
      unsupportedObjects: string;
      protectedSheet: string;
      groupOutsideRange: string;
      groupLimit: string;
      noGroups: string;
      noRoom: string;
    };
  };
  macSparkline: {
    title: string;
    dataRange: string;
    location: string;
    type: string;
    ok: string;
    cancel: string;
    invalidSource: string;
    locationOneCell: string;
    outOfBounds: string;
    chooseType: string;
    unavailable: string;
    protectedLocation: string;
  };
  macSlicer: {
    table: string;
    ok: string;
    cancel: string;
    noTable: string;
    noColumns: string;
    chooseTableAndColumn: string;
    insertFailed: string;
  };
  macWorkbookStats: {
    title: string;
    close: string;
    sheets: string;
    populatedCells: string;
    formulas: string;
    numbers: string;
    text: string;
    booleans: string;
    errors: string;
    tables: string;
    comments: string;
    hyperlinks: string;
    usedRows: string;
    usedColumns: string;
  };
  macInk: {
    strokeLabel: string;
    drawLabel: string;
  };
}
