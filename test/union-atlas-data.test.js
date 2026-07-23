const fs = require('fs');
const path = require('path');

const atlasPath = path.join(__dirname, '..', 'docs', 'union-atlas.html');
const atlas = fs.readFileSync(atlasPath, 'utf8');

describe('Union Atlas USUnions data bridge', () => {
  test('includes the exact seeded aggregate snapshot', () => {
    expect(atlas).toContain('surveyParticipation:62');
    expect(atlas).toContain('weeklyQuestionVotes:147');
    expect(atlas).toContain('{month:"Feb",total:878,newMembers:16,departed:7}');
    expect(atlas).toContain('avgCaseload:23.4');
    expect(atlas).toContain('submissionRate:71');
  });

  test('labels the source as demo rather than production', () => {
    expect(atlas).toContain('mode:"repo-demo"');
    expect(atlas).toContain('production:false');
    expect(atlas).toContain('Demo aggregate');
    expect(atlas).toContain('not production');
  });

  test('states and enforces the privacy boundary', () => {
    expect(atlas).toContain('Member names, roster rows or personal identifiers');
    expect(atlas).toContain('Grievance, discipline or case-level records');
    expect(atlas).toContain('Street addresses, steward names or authentication data');
    expect(atlas).not.toMatch(/[A-Z0-9._%+-]+@(?!example\.com)[A-Z0-9.-]+\.[A-Z]{2,}/i);
  });

  test('keeps the imported total separate from Atlas counts', () => {
    expect(atlas).toContain('keeps local operations separate from Atlas-wide organization and proximity counts');
    expect(atlas).toContain('Demo members');
  });
});
