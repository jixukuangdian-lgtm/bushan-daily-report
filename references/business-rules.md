# Business Rules

## Daily total

Daily totals are built from parsed platform entries for:

- Xiaohongshu
- Douyin
- Youzan
- WeChat Video big account
- WeChat Video small account

## Product-card rule

Product-card totals should mean:

- Xiaohongshu product card
- Douyin product card
- WeChat Video big-account product card
- WeChat Video small-account product card

## Xiaohongshu

- Live actual = live GMV - live refund
- Product-card actual = note GMV + merch-card GMV minus their matching refunds

## Douyin

- Live entry comes from `载体类型 = 直播`
- Product-card entry comes from `载体类型 = 商品卡 + 其他`
- Use the unrestricted time-range row when the sheet has a time-range dimension

## WeChat Video big account

- Preferred source: two `场景构成` files
- Total actual = `成交金额 - 成交退款金额`
- Live actual = `直播间成交金额 - 直播间成交退款金额`
- All non-live scenes count as product-card actual
- If the source structure changes, do not silently guess; record the mismatch and stop or downgrade safely

## WeChat Video small account

- Preferred source: one `场景构成` file
- Total actual = `成交金额 - 成交退款金额`
- Live actual = `直播间成交金额 - 直播间成交退款金额`
- All non-live scenes count as product-card actual

## Youzan

- GMV is summed by payment time from the order export
- Refund is summed by refund-complete time from the refund export
- Only successful refund records should count
- For late-month reporting, use a cumulative refund-complete export covering the full month through the report date
- Do not publish when the cumulative refund source is incomplete; stop instead of generating a wrong month refund total

## Month progress sheet

Displayed monthly rows:

- Xiaohongshu live
- Douyin live
- Product card
- WeChat Video big-account live
- WeChat Video small-account live
- Youzan
- Month total

`Month total` is the sum of the displayed rows, not a separate deduplicated business view.

## Rerun safety

- When rerunning an existing date, rewrite workbook rows for that date instead of skipping workbook writes
- Recompute month progress after the rewrite
- Then sync Base and resend only if needed
