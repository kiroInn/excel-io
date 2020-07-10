import { render, fireEvent } from '@testing-library/vue'
import Evaluation from '@/views/Evaluation.vue'

test('open config mappings on click', async () => {
  const { getByText } = render(Evaluation)

  const button = getByText('Config')
  await fireEvent.click(button)

  getByText('Config Variable')
})

test('edit config mappings on update', async () => {
    const { getByText, getByDisplayValue } = render(Evaluation)
    const button = getByText('Config')
    await fireEvent.click(button)
    getByText('Config Variable')

    const input1 = getByDisplayValue('CNCDU-106700:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}')
    await fireEvent.update(input1, 'CNCDU-106700:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}')
    getByDisplayValue('CNCDU-106700:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}');
    
  })