import type { CellValue } from '../../../engine/types.js';
import { yearFrac } from '../calendar.js';
import { numericResult, readNumber } from '../coercion.js';
import type { FormulaOperand } from '../types.js';
import type { FormulaReaderContext, NumericFunctionName } from './context.js';
import type { RangeReader } from './ranges.js';

/** Exactly the reader members this family touches. */
export type FinancialEvaluatorContext = Pick<FormulaReaderContext, 'readOperand'> &
  Pick<
    RangeReader,
    | 'formulaRangeArgBounds'
    | 'numericValuesInBounds'
    | 'numericValuesInFormulaRangeArg'
    | 'numericValuesInRangeWithShape'
  >;

/** Financial functions: depreciation, annuities (PMT/PV/FV/NPER/RATE and
 *  their per-period splits), discount securities, T-bills, rate conversions,
 *  and cash-flow schedules (NPV, MIRR, XNPV, IRR, XIRR, FVSCHEDULE). */
export function createFinancialEvaluator(ctx: FinancialEvaluatorContext) {
  const {
    readOperand,
    formulaRangeArgBounds,
    numericValuesInBounds,
    numericValuesInFormulaRangeArg,
    numericValuesInRangeWithShape,
  } = ctx;
  const financialFunction = (
    fn: NumericFunctionName,
    args: FormulaOperand[],
    rowOffset: number,
    colOffset: number,
  ): CellValue | undefined => {
    const financialType = (value: number): number | null => {
      const type = Math.trunc(value);
      return type === 0 || type === 1 ? type : null;
    };
    const annuityFactor = (rate: number, periods: number): number | null => {
      if (rate === 0) return periods;
      const factor = (1 + rate) ** periods;
      const value = (factor - 1) / rate;
      return Number.isFinite(value) ? value : null;
    };
    const financialFutureValue = (
      rate: number,
      periods: number,
      payment: number,
      presentValue: number,
      type: number,
    ): number => {
      if (rate === 0) return presentValue + payment * periods;
      const factor = (1 + rate) ** periods;
      return presentValue * factor + payment * (1 + rate * type) * ((factor - 1) / rate);
    };
    const financialPayment = (
      rate: number,
      periods: number,
      presentValue: number,
      futureValue: number,
      type: number,
    ): number | null => {
      if (periods === 0) return null;
      if (rate === 0) return -(presentValue + futureValue) / periods;
      const factor = (1 + rate) ** periods;
      const denominator = (1 + rate * type) * (factor - 1);
      if (denominator === 0) return null;
      const value = -((futureValue + presentValue * factor) * rate) / denominator;
      return Number.isFinite(value) ? value : null;
    };
    const depreciationInputs = (
      costOperand: FormulaOperand | undefined,
      salvageOperand: FormulaOperand | undefined,
      lifeOperand: FormulaOperand | undefined,
    ): { cost: number; salvage: number; life: number } | CellValue => {
      if (!costOperand || !salvageOperand || !lifeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const cost = readNumber(readOperand(costOperand, rowOffset, colOffset));
      const salvage = readNumber(readOperand(salvageOperand, rowOffset, colOffset));
      const life = readNumber(readOperand(lifeOperand, rowOffset, colOffset));
      if (cost === null || salvage === null || life === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (cost < 0 || salvage < 0 || life <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return { cost, salvage, life };
    };
    if (fn === 'SLN' || fn === 'SYD' || fn === 'DDB' || fn === 'DB') {
      const [costOperand, salvageOperand, lifeOperand, periodOperand, factorOperand] = args;
      const inputs = depreciationInputs(costOperand, salvageOperand, lifeOperand);
      if ('kind' in inputs) return inputs;
      const { cost, salvage, life } = inputs;
      if (fn === 'SLN') return numericResult((cost - salvage) / life);
      const rawPeriod = readNumber(
        readOperand(periodOperand as FormulaOperand, rowOffset, colOffset),
      );
      if (rawPeriod === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const period = Math.trunc(rawPeriod);
      if (period < 1 || (fn !== 'DB' && period > life)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'SYD') {
        return numericResult(((cost - salvage) * (life - period + 1) * 2) / (life * (life + 1)));
      }
      if (fn === 'DB') {
        const rawMonth =
          factorOperand === undefined
            ? 12
            : readNumber(readOperand(factorOperand, rowOffset, colOffset));
        if (rawMonth === null) return { kind: 'error', code: 15, text: '#VALUE!' };
        const month = Math.trunc(rawMonth);
        if (month < 1 || month > 12 || salvage > cost) {
          return { kind: 'error', code: 6, text: '#NUM!' };
        }
        const maxPeriod = month === 12 ? Math.trunc(life) : Math.trunc(life) + 1;
        if (period > maxPeriod) return { kind: 'error', code: 6, text: '#NUM!' };
        const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
        let accumulated = 0;
        let depreciation = 0;
        for (let currentPeriod = 1; currentPeriod <= period; currentPeriod += 1) {
          if (currentPeriod === 1) {
            depreciation = cost * rate * (month / 12);
          } else if (currentPeriod === Math.trunc(life) + 1) {
            depreciation = (cost - accumulated) * rate * ((12 - month) / 12);
          } else {
            depreciation = (cost - accumulated) * rate;
          }
          accumulated += depreciation;
        }
        return numericResult(depreciation);
      }
      const rawFactor =
        factorOperand === undefined
          ? 2
          : readNumber(readOperand(factorOperand, rowOffset, colOffset));
      if (rawFactor === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      if (rawFactor <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      let bookValue = cost;
      let depreciation = 0;
      for (let currentPeriod = 1; currentPeriod <= period; currentPeriod += 1) {
        depreciation = Math.min(bookValue * (rawFactor / life), Math.max(0, bookValue - salvage));
        bookValue -= depreciation;
      }
      return numericResult(depreciation);
    }
    if (fn === 'CUMIPMT' || fn === 'CUMPRINC') {
      const [
        rateOperand,
        periodsOperand,
        presentValueOperand,
        startOperand,
        endOperand,
        typeOperand,
      ] = args;
      if (
        !rateOperand ||
        !periodsOperand ||
        !presentValueOperand ||
        !startOperand ||
        !endOperand ||
        !typeOperand
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const periods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const presentValue = readNumber(readOperand(presentValueOperand, rowOffset, colOffset));
      const rawStart = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const rawEnd = readNumber(readOperand(endOperand, rowOffset, colOffset));
      const rawType = readNumber(readOperand(typeOperand, rowOffset, colOffset));
      if (
        rate === null ||
        periods === null ||
        presentValue === null ||
        rawStart === null ||
        rawEnd === null ||
        rawType === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const type = financialType(rawType);
      const startPeriod = Math.trunc(rawStart);
      const endPeriod = Math.trunc(rawEnd);
      if (
        type === null ||
        rate <= 0 ||
        periods <= 0 ||
        presentValue <= 0 ||
        startPeriod < 1 ||
        endPeriod < 1 ||
        startPeriod > endPeriod ||
        endPeriod > periods
      ) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const payment = financialPayment(rate, periods, presentValue, 0, type);
      if (payment === null) return { kind: 'error', code: 6, text: '#NUM!' };
      let balance = presentValue;
      let cumulative = 0;
      for (let currentPeriod = 1; currentPeriod <= endPeriod; currentPeriod += 1) {
        if (type === 1) balance += payment;
        const interest = currentPeriod === 1 && type === 1 ? 0 : -balance * rate;
        const principal = payment - interest;
        if (currentPeriod >= startPeriod) {
          cumulative += fn === 'CUMIPMT' ? interest : principal;
        }
        if (type === 0) balance += principal;
        else balance -= interest;
      }
      return numericResult(cumulative);
    }
    if (fn === 'ISPMT') {
      const [rateOperand, periodOperand, periodsOperand, presentValueOperand] = args;
      if (!rateOperand || !periodOperand || !periodsOperand || !presentValueOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const period = readNumber(readOperand(periodOperand, rowOffset, colOffset));
      const periods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const presentValue = readNumber(readOperand(presentValueOperand, rowOffset, colOffset));
      if (rate === null || period === null || periods === null || presentValue === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (periods === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult((-presentValue * rate * (periods - period)) / periods);
    }
    if (fn === 'EFFECT' || fn === 'NOMINAL') {
      const [rateOperand, periodsOperand] = args;
      if (!rateOperand || !periodsOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const rawPeriods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      if (rate === null || rawPeriods === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const periods = Math.trunc(rawPeriods);
      if (rate <= 0 || periods < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'EFFECT') return numericResult((1 + rate / periods) ** periods - 1);
      return numericResult(periods * ((1 + rate) ** (1 / periods) - 1));
    }
    if (fn === 'DOLLARDE' || fn === 'DOLLARFR') {
      const [dollarOperand, fractionOperand] = args;
      if (!dollarOperand || !fractionOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const dollar = readNumber(readOperand(dollarOperand, rowOffset, colOffset));
      const rawFraction = readNumber(readOperand(fractionOperand, rowOffset, colOffset));
      if (dollar === null || rawFraction === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const fraction = Math.trunc(rawFraction);
      if (fraction < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      const sign = Math.sign(dollar) || 1;
      const absolute = Math.abs(dollar);
      const integer = Math.trunc(absolute);
      const fractional = absolute - integer;
      if (fn === 'DOLLARDE') {
        return numericResult(sign * (integer + (fractional * 100) / fraction));
      }
      return numericResult(sign * (integer + (fractional * fraction) / 100));
    }
    if (fn === 'DISC' || fn === 'INTRATE' || fn === 'PRICEDISC' || fn === 'RECEIVED') {
      const [settlementOperand, maturityOperand, thirdOperand, redemptionOperand, basisOperand] =
        args;
      if (!settlementOperand || !maturityOperand || !thirdOperand || !redemptionOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const settlement = readNumber(readOperand(settlementOperand, rowOffset, colOffset));
      const maturity = readNumber(readOperand(maturityOperand, rowOffset, colOffset));
      const third = readNumber(readOperand(thirdOperand, rowOffset, colOffset));
      const redemption = readNumber(readOperand(redemptionOperand, rowOffset, colOffset));
      const basis =
        basisOperand === undefined
          ? 0
          : readNumber(readOperand(basisOperand, rowOffset, colOffset));
      if (
        settlement === null ||
        maturity === null ||
        third === null ||
        redemption === null ||
        basis === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (maturity <= settlement || third <= 0 || redemption <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const yearFraction = yearFrac(settlement, maturity, basis);
      if (yearFraction.kind !== 'number' || yearFraction.value <= 0) return yearFraction;
      if (fn === 'PRICEDISC') {
        return numericResult(redemption * (1 - third * yearFraction.value));
      }
      if (fn === 'RECEIVED') {
        const denominator = 1 - redemption * yearFraction.value;
        if (denominator === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
        return numericResult(third / denominator);
      }
      const denominator = fn === 'DISC' ? redemption : third;
      return numericResult((redemption - third) / denominator / yearFraction.value);
    }
    if (fn === 'ACCRINTM') {
      const [issueOperand, settlementOperand, rateOperand, parOperand, basisOperand] = args;
      if (!issueOperand || !settlementOperand || !rateOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const issue = readNumber(readOperand(issueOperand, rowOffset, colOffset));
      const settlement = readNumber(readOperand(settlementOperand, rowOffset, colOffset));
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const par =
        parOperand === undefined ? 1000 : readNumber(readOperand(parOperand, rowOffset, colOffset));
      const basis =
        basisOperand === undefined
          ? 0
          : readNumber(readOperand(basisOperand, rowOffset, colOffset));
      if (
        issue === null ||
        settlement === null ||
        rate === null ||
        par === null ||
        basis === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (settlement <= issue || rate <= 0 || par <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const yearFraction = yearFrac(issue, settlement, basis);
      if (yearFraction.kind !== 'number' || yearFraction.value <= 0) return yearFraction;
      return numericResult(par * rate * yearFraction.value);
    }
    if (fn === 'TBILLPRICE' || fn === 'TBILLYIELD' || fn === 'TBILLEQ') {
      const [settlementOperand, maturityOperand, thirdOperand] = args;
      if (!settlementOperand || !maturityOperand || !thirdOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const settlement = readNumber(readOperand(settlementOperand, rowOffset, colOffset));
      const maturity = readNumber(readOperand(maturityOperand, rowOffset, colOffset));
      const third = readNumber(readOperand(thirdOperand, rowOffset, colOffset));
      if (settlement === null || maturity === null || third === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const days = Math.trunc(maturity) - Math.trunc(settlement);
      if (days <= 0 || days > 365 || third <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'TBILLPRICE') {
        return numericResult(100 * (1 - (third * days) / 360));
      }
      if (fn === 'TBILLYIELD') {
        return numericResult(((100 - third) / third) * (360 / days));
      }
      const denominator = 360 - third * days;
      if (denominator <= 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult((365 * third) / denominator);
    }
    if (fn === 'RRI' || fn === 'PDURATION') {
      const [periodsOperand, presentValueOperand, futureValueOperand] = args;
      if (!periodsOperand || !presentValueOperand || !futureValueOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const first = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const presentValue = readNumber(readOperand(presentValueOperand, rowOffset, colOffset));
      const futureValue = readNumber(readOperand(futureValueOperand, rowOffset, colOffset));
      if (first === null || presentValue === null || futureValue === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (fn === 'RRI') {
        if (first <= 0 || presentValue <= 0 || futureValue <= 0) {
          return { kind: 'error', code: 6, text: '#NUM!' };
        }
        return numericResult((futureValue / presentValue) ** (1 / first) - 1);
      }
      const rate = first;
      if (rate <= 0 || presentValue <= 0 || futureValue <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return numericResult(Math.log(futureValue / presentValue) / Math.log(1 + rate));
    }
    if (fn === 'IPMT' || fn === 'PPMT') {
      const [
        rateOperand,
        periodOperand,
        periodsOperand,
        presentValueOperand,
        futureValueOperand,
        typeOperand,
      ] = args;
      if (!rateOperand || !periodOperand || !periodsOperand || !presentValueOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const rawPeriod = readNumber(readOperand(periodOperand, rowOffset, colOffset));
      const periods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const presentValue = readNumber(readOperand(presentValueOperand, rowOffset, colOffset));
      const futureValue =
        futureValueOperand === undefined
          ? 0
          : readNumber(readOperand(futureValueOperand, rowOffset, colOffset));
      const rawType =
        typeOperand === undefined ? 0 : readNumber(readOperand(typeOperand, rowOffset, colOffset));
      if (
        rate === null ||
        rawPeriod === null ||
        periods === null ||
        presentValue === null ||
        futureValue === null ||
        rawType === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const type = financialType(rawType);
      const period = Math.trunc(rawPeriod);
      if (type === null || period < 1 || period > periods || periods === 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const payment = financialPayment(rate, periods, presentValue, futureValue, type);
      if (payment === null) return { kind: 'error', code: 6, text: '#NUM!' };
      if (rate === 0) {
        const interest = 0;
        return { kind: 'number', value: fn === 'IPMT' ? interest : payment - interest };
      }
      let balance = presentValue;
      let interest = 0;
      for (let currentPeriod = 1; currentPeriod <= period; currentPeriod += 1) {
        if (type === 1) balance += payment;
        interest = currentPeriod === 1 && type === 1 ? 0 : -balance * rate;
        if (type === 0) balance += payment - interest;
        else balance -= interest;
      }
      return numericResult(fn === 'IPMT' ? interest : payment - interest);
    }
    if (fn === 'PMT' || fn === 'PV' || fn === 'FV' || fn === 'NPER' || fn === 'RATE') {
      const [rateOperand, periodsOperand, thirdOperand, fourthOperand, typeOperand, guessOperand] =
        args;
      if (!rateOperand || !periodsOperand || !thirdOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const periods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const third = readNumber(readOperand(thirdOperand, rowOffset, colOffset));
      const fourth =
        fourthOperand === undefined
          ? 0
          : readNumber(readOperand(fourthOperand, rowOffset, colOffset));
      const rawType =
        typeOperand === undefined ? 0 : readNumber(readOperand(typeOperand, rowOffset, colOffset));
      const guess =
        guessOperand === undefined
          ? 0.1
          : readNumber(readOperand(guessOperand, rowOffset, colOffset));
      if (
        rate === null ||
        periods === null ||
        third === null ||
        fourth === null ||
        rawType === null ||
        guess === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const type = financialType(rawType);
      if (type === null || (fn === 'RATE' ? rate === 0 : periods === 0)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'RATE') {
        const payment = periods;
        const presentValue = third;
        const futureValue = fourth;
        const zeroRateValue =
          financialFutureValue(0, rate, payment, presentValue, type) + futureValue;
        if (Math.abs(zeroRateValue) < 1e-9) return { kind: 'number', value: 0 };
        let current = guess;
        for (let iteration = 0; iteration < 50; iteration += 1) {
          if (current <= -1) return { kind: 'error', code: 6, text: '#NUM!' };
          const value =
            financialFutureValue(current, rate, payment, presentValue, type) + futureValue;
          if (!Number.isFinite(value)) return { kind: 'error', code: 6, text: '#NUM!' };
          if (Math.abs(value) < 1e-9) return numericResult(current);
          const step = Math.max(Math.abs(current) * 1e-6, 1e-7);
          const high =
            financialFutureValue(current + step, rate, payment, presentValue, type) + futureValue;
          const low =
            financialFutureValue(current - step, rate, payment, presentValue, type) + futureValue;
          const derivative = (high - low) / (2 * step);
          if (!Number.isFinite(derivative) || derivative === 0) {
            return { kind: 'error', code: 6, text: '#NUM!' };
          }
          const next = current - value / derivative;
          if (!Number.isFinite(next)) return { kind: 'error', code: 6, text: '#NUM!' };
          if (Math.abs(next - current) < 1e-10) return numericResult(next);
          current = next;
        }
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'NPER') {
        const payment = periods;
        const presentValue = third;
        const futureValue = fourth;
        if (rate === 0) {
          if (payment === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
          return numericResult(-(presentValue + futureValue) / payment);
        }
        const adjustedPayment = payment * (1 + rate * type);
        const numerator = adjustedPayment - futureValue * rate;
        const denominator = adjustedPayment + presentValue * rate;
        const ratio = numerator / denominator;
        if (ratio <= 0 || rate <= -1) {
          return { kind: 'error', code: 6, text: '#NUM!' };
        }
        return numericResult(Math.log(ratio) / Math.log(1 + rate));
      }
      if (fn === 'PMT') {
        const presentValue = third;
        const futureValue = fourth;
        if (rate === 0) return numericResult(-(presentValue + futureValue) / periods);
        const factor = (1 + rate) ** periods;
        const denominator = (1 + rate * type) * (factor - 1);
        if (denominator === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
        return numericResult(-((futureValue + presentValue * factor) * rate) / denominator);
      }
      const payment = third;
      const factor = (1 + rate) ** periods;
      const annuity = annuityFactor(rate, periods);
      if (annuity === null) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'FV') {
        const presentValue = fourth;
        return numericResult(-(presentValue * factor + payment * (1 + rate * type) * annuity));
      }
      const futureValue = fourth;
      if (factor === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(-(futureValue + payment * (1 + rate * type) * annuity) / factor);
    }
    return undefined;
  };
  const npvValue = (
    operand: Extract<FormulaOperand, { kind: 'npv' }>,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const rate = readNumber(readOperand(operand.rate, rowOffset, colOffset));
    if (rate === null || rate === -1) {
      return {
        kind: 'error',
        code: rate === -1 ? 1 : 15,
        text: rate === -1 ? '#DIV/0!' : '#VALUE!',
      };
    }
    let period = 1;
    let result = 0;
    for (const arg of operand.values) {
      if (arg.kind === 'range' || arg.kind === 'dynamic-range') {
        const bounds = formulaRangeArgBounds(arg, rowOffset, colOffset);
        const values = bounds ? numericValuesInBounds(bounds) : null;
        if (values === null) return { kind: 'error', code: 15, text: '#VALUE!' };
        for (const value of values) {
          result += value / (1 + rate) ** period;
          period += 1;
        }
        continue;
      }
      const value = readNumber(readOperand(arg.operand, rowOffset, colOffset));
      if (value === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      result += value / (1 + rate) ** period;
      period += 1;
    }
    return Number.isFinite(result)
      ? { kind: 'number', value: result }
      : { kind: 'error', code: 6, text: '#NUM!' };
  };
  const mirrValue = (
    operand: Extract<FormulaOperand, { kind: 'mirr' }>,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const values = numericValuesInFormulaRangeArg(operand.values, rowOffset, colOffset);
    const financeRate = readNumber(readOperand(operand.financeRate, rowOffset, colOffset));
    const reinvestRate = readNumber(readOperand(operand.reinvestRate, rowOffset, colOffset));
    if (values === null || financeRate === null || reinvestRate === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (values.length < 2 || financeRate === -1 || reinvestRate === -1) {
      return {
        kind: 'error',
        code: financeRate === -1 || reinvestRate === -1 ? 1 : 6,
        text: financeRate === -1 || reinvestRate === -1 ? '#DIV/0!' : '#NUM!',
      };
    }
    let presentValueNegative = 0;
    let futureValuePositive = 0;
    for (let index = 0; index < values.length; index += 1) {
      const value = values[index] ?? 0;
      if (value < 0) presentValueNegative += value / (1 + financeRate) ** index;
      else if (value > 0)
        futureValuePositive += value * (1 + reinvestRate) ** (values.length - 1 - index);
    }
    if (presentValueNegative === 0 || futureValuePositive === 0) {
      return { kind: 'error', code: 1, text: '#DIV/0!' };
    }
    const result = (-futureValuePositive / presentValueNegative) ** (1 / (values.length - 1)) - 1;
    return Number.isFinite(result)
      ? { kind: 'number', value: result }
      : { kind: 'error', code: 6, text: '#NUM!' };
  };
  const xnpvValue = (
    operand: Extract<FormulaOperand, { kind: 'xnpv' }>,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const rate = readNumber(readOperand(operand.rate, rowOffset, colOffset));
    const values = numericValuesInRangeWithShape(operand.values, rowOffset, colOffset);
    const dates = numericValuesInRangeWithShape(operand.dates, rowOffset, colOffset);
    if (rate === null || values === null || dates === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (
      values.width !== dates.width ||
      values.height !== dates.height ||
      values.values.length === 0 ||
      rate === -1
    ) {
      return {
        kind: 'error',
        code: rate === -1 ? 1 : 15,
        text: rate === -1 ? '#DIV/0!' : '#VALUE!',
      };
    }
    const firstDate = dates.values[0] as number;
    let result = 0;
    for (let index = 0; index < values.values.length; index += 1) {
      const value = values.values[index] as number;
      const date = dates.values[index] as number;
      if (date < firstDate) return { kind: 'error', code: 6, text: '#NUM!' };
      result += value / (1 + rate) ** ((date - firstDate) / 365);
    }
    return Number.isFinite(result)
      ? { kind: 'number', value: result }
      : { kind: 'error', code: 6, text: '#NUM!' };
  };
  const irrValue = (
    operand: Extract<FormulaOperand, { kind: 'irr' }>,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const values = numericValuesInFormulaRangeArg(operand.values, rowOffset, colOffset);
    const rawGuess = operand.guess
      ? readNumber(readOperand(operand.guess, rowOffset, colOffset))
      : 0.1;
    if (values === null || rawGuess === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (values.length === 0 || rawGuess <= -1) return { kind: 'error', code: 6, text: '#NUM!' };
    let hasPositive = false;
    let hasNegative = false;
    for (const value of values) {
      if (value > 0) hasPositive = true;
      if (value < 0) hasNegative = true;
    }
    if (!hasPositive || !hasNegative) return { kind: 'error', code: 1, text: '#DIV/0!' };
    const evaluatePeriodicNpv = (rate: number): number => {
      let total = 0;
      for (let index = 0; index < values.length; index += 1) {
        total += (values[index] as number) / (1 + rate) ** index;
      }
      return total;
    };
    let current = rawGuess;
    for (let iteration = 0; iteration < 100; iteration += 1) {
      if (current <= -1) return { kind: 'error', code: 6, text: '#NUM!' };
      const value = evaluatePeriodicNpv(current);
      if (!Number.isFinite(value)) return { kind: 'error', code: 6, text: '#NUM!' };
      if (Math.abs(value) < 1e-7) return { kind: 'number', value: current };
      const step = Math.max(Math.abs(current) * 1e-6, 1e-7);
      const high = evaluatePeriodicNpv(current + step);
      const low = evaluatePeriodicNpv(current - step);
      const derivative = (high - low) / (2 * step);
      if (!Number.isFinite(derivative) || derivative === 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const next = current - value / derivative;
      if (!Number.isFinite(next)) return { kind: 'error', code: 6, text: '#NUM!' };
      if (Math.abs(next - current) < 1e-10) return { kind: 'number', value: next };
      current = next;
    }
    return { kind: 'error', code: 6, text: '#NUM!' };
  };
  const xirrValue = (
    operand: Extract<FormulaOperand, { kind: 'xirr' }>,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const values = numericValuesInRangeWithShape(operand.values, rowOffset, colOffset);
    const dates = numericValuesInRangeWithShape(operand.dates, rowOffset, colOffset);
    const rawGuess = operand.guess
      ? readNumber(readOperand(operand.guess, rowOffset, colOffset))
      : 0.1;
    if (values === null || dates === null || rawGuess === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (
      values.width !== dates.width ||
      values.height !== dates.height ||
      values.values.length === 0 ||
      rawGuess <= -1
    ) {
      return {
        kind: 'error',
        code: rawGuess <= -1 ? 6 : 15,
        text: rawGuess <= -1 ? '#NUM!' : '#VALUE!',
      };
    }
    const firstDate = dates.values[0] as number;
    let hasPositive = false;
    let hasNegative = false;
    for (let index = 0; index < values.values.length; index += 1) {
      const value = values.values[index] as number;
      const date = dates.values[index] as number;
      if (date < firstDate) return { kind: 'error', code: 6, text: '#NUM!' };
      if (value > 0) hasPositive = true;
      if (value < 0) hasNegative = true;
    }
    if (!hasPositive || !hasNegative) return { kind: 'error', code: 1, text: '#DIV/0!' };
    const evaluateXnpv = (rate: number): number => {
      let total = 0;
      for (let index = 0; index < values.values.length; index += 1) {
        total +=
          (values.values[index] as number) /
          (1 + rate) ** (((dates.values[index] as number) - firstDate) / 365);
      }
      return total;
    };
    let current = rawGuess;
    for (let iteration = 0; iteration < 100; iteration += 1) {
      if (current <= -1) return { kind: 'error', code: 6, text: '#NUM!' };
      const value = evaluateXnpv(current);
      if (!Number.isFinite(value)) return { kind: 'error', code: 6, text: '#NUM!' };
      if (Math.abs(value) < 1e-7) return { kind: 'number', value: current };
      const step = Math.max(Math.abs(current) * 1e-6, 1e-7);
      const high = evaluateXnpv(current + step);
      const low = evaluateXnpv(current - step);
      const derivative = (high - low) / (2 * step);
      if (!Number.isFinite(derivative) || derivative === 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const next = current - value / derivative;
      if (!Number.isFinite(next)) return { kind: 'error', code: 6, text: '#NUM!' };
      if (Math.abs(next - current) < 1e-10) return { kind: 'number', value: next };
      current = next;
    }
    return { kind: 'error', code: 6, text: '#NUM!' };
  };
  const fvScheduleValue = (
    operand: Extract<FormulaOperand, { kind: 'fv-schedule' }>,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const principal = readNumber(readOperand(operand.principal, rowOffset, colOffset));
    const schedule = numericValuesInFormulaRangeArg(operand.schedule, rowOffset, colOffset);
    if (principal === null || schedule === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let value = principal;
    for (const rate of schedule) value *= 1 + rate;
    return Number.isFinite(value)
      ? { kind: 'number', value }
      : { kind: 'error', code: 6, text: '#NUM!' };
  };
  return {
    financialFunction,
    npvValue,
    mirrValue,
    xnpvValue,
    irrValue,
    xirrValue,
    fvScheduleValue,
  };
}
