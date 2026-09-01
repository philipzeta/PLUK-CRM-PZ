import { num } from './utils.js';

// Shared calculation logic that mirrors the original workbook's formulas,
// so tabs that reference each other (ACTION PLAN reads AGAPE's annual
// premium target, SCORECARD reads monthly APE) stay in sync.

export function computeAgape(d) {
  const groupTotal = (items) => items.reduce((s, it) => s + num(it.amount), 0);
  const sections = {};
  d.sections.forEach((sec) => {
    sections[sec.id] = sec.groups.reduce((s, g) => s + groupTotal(g.items), 0);
  });
  const goalsTotal = groupTotal(d.goals);
  const ftJobTotal = groupTotal(d.ftJobItems);

  const livingTotal = sections.living || 0;
  const businessTotal = sections.business || 0;
  const savingsTotal = sections.savings || 0;

  const incomeRequirement = livingTotal + businessTotal + savingsTotal + goalsTotal;
  const incomeFromEmployment = num(d.incomeFromEmployment);
  const targetFromBusiness = incomeRequirement - incomeFromEmployment;
  const annualIncomeRequirement = targetFromBusiness * 12;
  const avgFYC = num(d.avgFirstYearCommission, 0.3) || 0.0001;
  const annualPremiumTarget = annualIncomeRequirement / avgFYC;
  const avgApe = num(d.avgApePerCase, 1) || 1;
  const numCases = annualPremiumTarget / avgApe;

  const prospects = numCases * 15;
  const appointments = numCases * 5;
  const meetings = numCases * 3;
  const closed = numCases * 1;

  return {
    livingTotal, businessTotal, savingsTotal, goalsTotal, ftJobTotal,
    incomeRequirement, targetFromBusiness, annualIncomeRequirement,
    annualPremiumTarget, numCases,
    prospects, appointments, meetings, closed,
  };
}

// ACTION PLAN monthly cascade. Returns arrays of length 12.
export function computeActionPlan(ap, agapeAnnualPremiumTarget) {
  const targetApe = Array(12).fill((agapeAnnualPremiumTarget || 0) / 12);
  const avgCase = ap.sales.averageCaseSize.map((v) => num(v));
  const numCases = ap.sales.numberOfCases.map((v) => num(v));
  const actualApe = avgCase.map((v, i) => v * numCases[i]);
  const income = actualApe.map((v) => v * 0.35);

  const excess = [];
  for (let i = 0; i < 12; i++) {
    const prior = i === 0 ? 0 : excess[i - 1];
    excess.push(targetApe[i] - actualApe[i] + (i === 0 ? 0 : prior));
  }

  const beginMP = [];
  const endMP = [];
  const targetRecruits = ap.recruitment.targetNewRecruits.map((v) => num(v));
  const attrition = ap.recruitment.provisionForAttrition.map((v) => num(v));
  for (let i = 0; i < 12; i++) {
    beginMP.push(i === 0 ? num(ap.recruitment.beginningManpowerJan) : endMP[i - 1]);
    endMP.push(beginMP[i] + targetRecruits[i] - attrition[i]);
  }

  const targetCases = ap.salesActivity.targetCasesForMonth.map((v) => num(v));
  const prospects15 = targetCases.map((v) => v * 15);
  const appts5 = targetCases.map((v) => v * 5);
  const meetings3 = targetCases.map((v) => v * 3);

  const targetCoded = ap.recruitmentActivity.targetCoded.map((v) => num(v));
  const showUps = targetCoded.map((v) => v * 5);
  const confirmedBYB = showUps.map((v) => v * 2);
  const recProspects = confirmedBYB.map((v) => v * 5);

  return {
    targetApe, avgCase, numCases, actualApe, income, excess,
    beginMP, targetRecruits, attrition, endMP,
    targetCases, prospects15, appts5, meetings3,
    targetCoded, showUps, confirmedBYB, recProspects,
  };
}
