import { render, fireEvent, cleanup } from '@testing-library/vue'
import App from '@/App.vue'
import Split from '@/views/Split.vue'
import Evaluation from '@/views/Evaluation.vue'
const routes = [
    {
        path: "/",
        redirect: "/split"
      },
      {
        path: "/split",
        name: "split",
        component: Split
      },
      {
        path: "/evaluation",
        name: "Evaluation",
        component: Evaluation
      }
  ]

test('change function page on click', async () => {
  // const { getByText } = render(App,{routes});

  // getByText('Choose a file');
  // const label = getByText('evaluation')
  // await fireEvent.click(label)

  // getByText('Choose multiple files')
})