import { nextTick } from 'vue'
import { driver, type DriveStep } from 'driver.js'
import 'driver.js/dist/driver.css'
import { useAppConfigStore } from '@/stores/appConfigStore'

type Translate = (key: string) => string
type Screen = 'table' | 'form'

const STEPS: { id?: string; screen?: Screen }[] = [
  {},
  { id: 'screen' },
  { id: 'open' },
  { id: 'menu' },
  { id: 'export' },
  { id: 'undo' },
  { id: 'find' },
  { id: 'table', screen: 'table' },
  { id: 'row', screen: 'table' },
  { id: 'sum', screen: 'table' },
  { id: 'form', screen: 'form' },
  { id: 'form-nav', screen: 'form' },
  { id: 'form-fields', screen: 'form' },
  { id: 'form-actions', screen: 'form' },
  { id: 'form-add', screen: 'form' },
  { id: 'sheets' },
  { id: 'status' },
  { id: 'home' },
]

function matches(id: string): HTMLElement[] {
  return [...document.querySelectorAll(`[data-tour="${id}"]`)].filter((el): el is HTMLElement => el instanceof HTMLElement)
}

function visible(el: HTMLElement): boolean {
  const box = el.getBoundingClientRect()
  return box.width > 0 && box.height > 0
}

function shown(id: string): boolean {
  return matches(id).some(visible)
}

function inPage(id: string): boolean {
  return document.querySelector(`[data-tour="${id}"]`) instanceof HTMLElement
}

function reveal(step?: { screen?: Screen }) {
  if (!step?.screen || window.innerWidth >= 768) return
  useAppConfigStore().setPhoneScreen(step.screen)
}

let active: ReturnType<typeof driver> | null = null

export function startPageTour(t: Translate) {
  active?.destroy()
  const defs = STEPS.filter((step) => !step.id || shown(step.id) || (step.screen && inPage(step.id)))
  const steps: DriveStep[] = defs.map((step) => ({
    element: step.id
      ? () => matches(step.id!).find(visible) ?? matches(step.id!)[0]
      : undefined,
    popover: {
      title: t(step.id ? `tour.${step.id}.title` : 'tour.welcome.title'),
      description: t(step.id ? `tour.${step.id}.body` : 'tour.welcome.body'),
    },
  }))
  const tour = driver({
    showProgress: true,
    progressText: '{{current}} / {{total}}',
    nextBtnText: t('tour.next'),
    prevBtnText: t('tour.prev'),
    doneBtnText: t('tour.done'),
    skipMissingElement: true,
    steps,
    onNextClick: (_el, _step, opts) => {
      reveal(defs[(opts.state.activeIndex ?? 0) + 1])
      nextTick(() => nextTick(() => opts.driver.moveNext()))
    },
    onPrevClick: (_el, _step, opts) => {
      reveal(defs[(opts.state.activeIndex ?? 0) - 1])
      nextTick(() => nextTick(() => opts.driver.movePrevious()))
    },
  })
  active = tour
  tour.drive()
}
