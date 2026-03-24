## Summary
<!-- 1-3 bullet points describing what this PR does and why -->

-

## Error Handling Checklist

<!-- Check all that apply. Leave unchecked items that don't apply to your change. -->

- [ ] New server functions return `{ success, message }` response envelope
- [ ] New server functions accept `sessionToken` and call `_resolveCallerEmail()` or `_requireStewardAuth()`
- [ ] Error paths return safe fallback values (not thrown exceptions) to the client
- [ ] New function added to `test/auth-denial.test.js` coverage
- [ ] Error scenarios tested in module test or `test/negative-paths.test.js`
- [ ] No new OAuth scopes introduced (or re-auth steps documented below)
- [ ] `npm run test:guards` passes locally

## Test Plan
<!-- How did you verify this change works? -->

- [ ] Unit tests pass (`npm run test:unit`)
- [ ] Deploy guards pass (`npm run test:guards`)
- [ ] Manual testing in dev deployment
