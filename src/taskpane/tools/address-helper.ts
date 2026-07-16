/**
 * address-helper.ts - ТОЧКА ЕДИНОГО ВХОДА для получения Excel.Range из строки.
 *
 * ВАЖНО: валидирует адрес через _shared/validation (path traversal защита)
 * ПЕРЕД делегированием в resolveRange из _shared/address.
 *
 * Все инструменты, которым нужен Excel.Range из строки адреса, обязаны
 * использовать getRangeSafe (или _shared/address.resolveRange напрямую с
 * явной предварительной валидацией).
 *
 * См. docs/03-TOOLS-SPEC.md §3.2 (валидация аргументов обязательна).
 */

import { resolveRange } from "./_shared/address";
import { assertFullAddress } from "./_shared/validation";

/**
 * Получить Excel.Range из строки адреса с обязательной валидацией.
 *
 * Принимает: "A1", "A1:B2", "Лист!A1", "'Мой лист'!A1:B2".
 * Бросает Error для невалидных/опасных адресов (path traversal, etc).
 *
 * Не делает context.sync() - это ответственность вызывающего (батчинг).
 */
export function getRangeSafe(
  context: Excel.RequestContext,
  address: string,
): Excel.Range {
  assertFullAddress(address);
  return resolveRange(context, address);
}
